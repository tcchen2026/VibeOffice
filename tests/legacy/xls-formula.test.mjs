import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { load } from '../../tools/legacy/load.mjs';
import { words, dword, rec, str, workbook, cfb } from './xls-fixtures.mjs';
const B = load(),
  X = B.xls;
if (process.env.XLS_C2_BEFORE) {
  for (const key of ['formula', 'prepareFormulas', 'formulaConstant']) delete X[key];
  vm.runInThisContext(fs.readFileSync(process.env.XLS_C2_BEFORE + '/xls-write.js', 'utf8'));
}
if (process.env.XLS_C2_REVIEW_BEFORE)
  for (const file of ['xls-formula', 'xls-formula-book'])
    vm.runInThisContext(fs.readFileSync(process.env.XLS_C2_REVIEW_BEFORE + '/' + file + '.js', 'utf8'));
if (process.env.XLS_C3_NAMES_BEFORE)
  for (const file of ['xls-formula', 'xls-formula-book'])
    vm.runInThisContext(fs.readFileSync(process.env.XLS_C3_NAMES_BEFORE + '/' + file + '.js', 'utf8'));
const cat = (...parts) => Buffer.concat(parts.map(p => (Buffer.isBuffer(p) ? p : Buffer.from(p))));
const integer = n => cat([0x1e], words(n));
const ref = (row, col, id = 0x24) => cat([id], words(row, col));
const number = n => {
  const b = Buffer.alloc(9);
  b[0] = 1;
  b.writeDoubleLE(n, 1);
  return b;
};
function formula(row, col, tokens, value = 0, extra = Buffer.alloc(0), type = 'n') {
  const b = Buffer.alloc(22);
  b.writeUInt16LE(row);
  b.writeUInt16LE(col, 2);
  b.writeUInt16LE(tokens.length, 20);
  if (type === 'n') b.writeDoubleLE(value, 6);
  else {
    b.writeUInt16LE(65535, 12);
    b[6] = { s: 0, b: 1, e: 2, empty: 3 }[type];
    b[8] = +value || 0;
  }
  return rec(6, cat(b, tokens, extra));
}
function defined(name, tokens, scope = -1) {
  const b = Buffer.alloc(14);
  b[3] = name.length;
  b.writeUInt16LE(tokens.length, 4);
  b.writeUInt16LE(scope + 1, 8);
  return rec(0x18, cat(b, [1], Buffer.from(name, 'utf16le'), tokens));
}
function prepare(globals, records) {
  const book = X.read(workbook(globals, [{ name: 'Data', records }]), { rich: true }),
    losses = [];
  return { book, losses, result: X.prepareFormulas(book, losses) };
}
async function converted(globals, records) {
  const result = await X.convert(cfb(workbook(globals, [{ name: 'Data', records }]))),
    z = await L.zip.read(result.pkg);
  return { result, z };
}

test('BIFF RPN preserves operand order, strings, optional arguments and optimized SUM', () => {
  assert.equal(X.formula(cat(integer(2), integer(3), [4], integer(5), [5])), '(2-3)*5');
  assert.equal(X.formula(cat([0x17, 3, 0], Buffer.from('a"b'))), '"a""b"');
  assert.equal(X.formula(cat(integer(1), [0x16], integer(7), [0x22, 3], words(1))), 'IF(1,,7)');
  assert.equal(X.formula(cat(ref(0, 0xc000), [0x19, 0x10, 0, 0])), 'SUM(A1)');
});

test('ordinary coordinates differ from shared/name offsets and preserve absolute axes', () => {
  const tokens = ref(65535, 0xc0ff);
  assert.equal(X.formula(tokens, undefined, { row: 4, col: 3 }), 'IV65536');
  assert.equal(X.formula(tokens, undefined, { row: 4, col: 3, shared: true }), 'C4');
  assert.equal(X.formula(ref(7, 0x4002), undefined, { row: 4, col: 3, shared: true }), 'F$8');
  assert.equal(X.formula(ref(65535, 0x8002, 0x2c), undefined, { row: 4, col: 3 }), '$C4');
  assert.equal(X.formula(cat([0x25], words(0, 65535, 0, 1))), '$A:$B');
});

test('defined names retain current-sheet relative references with the required bang prefix', async () => {
  const { z } = await converted(
    [
      defined('CurrentValue', ref(199, 0xc015)),
      defined('CurrentArea', cat([0x25], words(0, 1, 0xc000, 0xc001)))
    ],
    []
  );
  const book = await z.get('xl/workbook.xml').text();
  assert.match(book, /<definedName name="CurrentValue">!V200<\/definedName>/);
  assert.match(book, /<definedName name="CurrentArea">!A1:B2<\/definedName>/);
  assert.equal(X.formula(ref(199, 0xc015)), 'V200');
});

test('invalid and duplicate names are isolated without rebinding dependent cell formulas', () => {
  const bad = ['?Bad', 'A1', 'XFD1048576', 'C1244_suffix', 'Name¦', 'Name©', 'Name»', 'Name\x8f'],
    good = ['_?', 'Données', 'Nameㅁ', 'XFE1', 'A1048577', '_€', '±Bad', '￥', 'Name¶', 'Name§'];
  const names = [...bad, ...good].map(n => defined(n, integer(2)));
  names.push(defined('Données', integer(3)), defined('Données', integer(4), 0));
  const { book, result, losses } = prepare(names, [
    formula(0, 0, cat([0x23], words(1, 0)), 42),
    formula(0, 1, cat([0x23], words(bad.length + 1, 0)), 2)
  ]);
  for (const n of bad) assert.ok(!result.defined.includes('name="' + n + '"'));
  for (const n of good) assert.ok(result.defined.includes('name="' + n + '"'));
  assert.match(result.defined, /name="Données" localSheetId="0">4/);
  assert.equal((result.defined.match(/name="Données"/g) || []).length, 2);
  assert.equal(book.sheets[0].cells[0].value, 42);
  assert.equal(book.sheets[0].cells[0].formulaOut, undefined);
  assert.equal(book.sheets[0].cells[1].formulaOut.text, '_?');
  assert.equal(result.stats.dropped, 1);
  assert.equal(losses.filter(l => l.id === 'legacy-defined-name').length, bad.length + 1);
});

test('formula array constants consume trailing data after memory extensions', () => {
  const tokens = cat([0x26], Buffer.alloc(6), [0x20], Buffer.alloc(7));
  const extra = cat(words(1), Buffer.alloc(8), [1], words(1), number(1), number(2), number(3), number(4));
  assert.equal(X.formula(tokens, extra), '{1,2;3,4}');
  assert.throws(() => X.formula(tokens, extra.subarray(0, -1)), { name: 'LegacyError' });
  assert.throws(() => X.formula(Buffer.from([3])), { name: 'LegacyError' });
  assert.throws(() => X.formula(cat(integer(1), integer(2))), { name: 'LegacyError' });
  assert.throws(() => X.formula(Buffer.from([0x18, 0xff])), { name: 'LegacyError' });
});

test('shared, array and data-table groups retain owners and member cached values', () => {
  const sharedTokens = cat(ref(0, 0xc0ff, 0x2c), integer(2), [5]);
  const shared = rec(0x4bc, cat(words(0, 1), [1, 1], words(0, sharedTokens.length), sharedTokens));
  const arrayTokens = cat([0x20], Buffer.alloc(7)),
    arrayExtra = cat([1], words(0), number(10), number(20));
  const array = rec(
    0x221,
    cat(words(3, 3), [0, 1], words(0), dword(0), words(arrayTokens.length), arrayTokens, arrayExtra)
  );
  const table = rec(0x236, cat(words(5, 6), [1, 2], words(8, 1, 4, 2, 5)));
  const records = [
    formula(0, 1, cat([1], words(0, 1)), 4),
    shared,
    formula(1, 1, cat([1], words(0, 1)), 8),
    formula(3, 0, cat([1], words(3, 0)), 10),
    array,
    formula(3, 1, cat([1], words(3, 0)), 20),
    formula(5, 1, cat([2], words(5, 1)), 1),
    table,
    formula(6, 2, cat([2], words(5, 1)), 2)
  ];
  const { book, result, losses } = prepare([], records),
    cells = book.sheets[0].cells;
  assert.equal(losses.length, 0);
  assert.equal(result.stats.converted, 6);
  assert.equal(cells[0].formulaOut.text, 'A1*2');
  assert.equal(cells[1].formulaOut.text, 'A2*2');
  assert.deepEqual(cells[2].formulaOut.attrs, { t: 'array', ref: 'A4:B4', aca: null });
  assert.equal(cells[3].formulaOut.member, true);
  assert.equal(cells[4].formulaOut.attrs.r1, 'E2');
  assert.equal(cells[4].formulaOut.attrs.r2, 'F3');
  assert.equal(cells[4].formulaOut.attrs.dt2D, 1);
  assert.equal(cells[5].formulaOut.member, true);
  assert.deepEqual(
    cells.map(c => c.value),
    [4, 8, 10, 20, 1, 2]
  );
});

test('a damaged formula loses only its formula and reports the cell; other formulas survive', () => {
  const { book, result, losses } = prepare(
    [],
    [formula(0, 0, integer(7), 7), formula(0, 1, Buffer.from([3]), 42)]
  );
  assert.equal(result.stats.converted, 1);
  assert.equal(result.stats.dropped, 1);
  assert.equal(book.sheets[0].cells[1].value, 42);
  assert.equal(losses.length, 1);
  assert.equal(losses[0].notify, true);
  assert.deepEqual(losses[0].detail.examples, ['B1']);
  assert.equal(losses[0].place, 'Data');
});

test('scoped names, external cached cells and named functions keep their workbook bindings', async () => {
  const globals = [
    defined('LocalValue', integer(9), 0),
    rec(0x1ae, words(1, 0x401)),
    rec(0x1ae, cat(words(1), str('\x01\x05\x10https://a/x.xlsx'), str('Elsewhere'))),
    rec(0x17, cat(words(2), words(0, 0, 0), words(1, 0, 0))),
    rec(0x59, words(1, 0)),
    rec(0x5a, cat([0, 0], words(0), number(12)))
  ];
  const { result, z } = await converted(globals, [
    formula(0, 0, cat([0x23], words(1, 0)), 9),
    formula(0, 1, cat([0x3a], words(1, 0, 0)), 12)
  ]);
  assert.equal(result.stats.formulas.converted, 2);
  const book = await z.get('xl/workbook.xml').text(),
    external = await z.get('xl/externalLinks/externalLink1.xml').text();
  assert.match(book, /definedName name="LocalValue" localSheetId="0">9/);
  assert.doesNotMatch(book, /date1904=/);
  assert.match(external, /<cell r="A1"><v>12<\/v>/);
  assert.match(await z.get('xl/externalLinks/_rels/externalLink1.xml.rels').text(), /https:\/\/a\/x.xlsx/);
  const named = cat([0x39], words(1, 1, 0), integer(2), [0x22, 2], words(255));
  assert.equal(
    X.formula(named, undefined, { externalName: () => ({ text: '[1]!CustomFn', name: '[1]!CustomFn' }) }),
    '[1]!CustomFn(2)'
  );
});

test('converted string, empty, Boolean and error formula caches have formula-cell types', async () => {
  const records = [
    formula(0, 0, cat([0x17, 1, 0, 65]), 0, undefined, 's'),
    rec(0x207, str('A')),
    formula(0, 1, cat([0x17, 0, 0]), 0, undefined, 'empty'),
    formula(0, 2, cat([0x1d, 1]), 1, undefined, 'b'),
    formula(0, 3, cat([0x1c, 7]), 7, undefined, 'e')
  ];
  const { z } = await converted([], records),
    xml = await z.get('xl/worksheets/sheet1.xml').text();
  assert.match(xml, /<c r="A1" t="str"><f>&quot;A&quot;<\/f><v>A<\/v>/);
  assert.match(xml, /<c r="B1" t="str"><f>&quot;&quot;<\/f><v><\/v>/);
  assert.match(xml, /<c r="C1" t="b"><f>TRUE<\/f><v>1<\/v>/);
  assert.match(xml, /<c r="D1" t="e"><f>#DIV\/0!<\/f><v>#DIV\/0!<\/v>/);
});

test('worksheet calculation settings retain manual mode, iteration and reference style', async () => {
  const delta = Buffer.alloc(8);
  delta.writeDoubleLE(0.125);
  const { z } = await converted(
    [rec(0xe, words(0))],
    [
      rec(0xd, words(0)),
      rec(0xc, words(200)),
      rec(0xf, words(0)),
      rec(0x10, delta),
      rec(0x11, words(1)),
      rec(0x5f, words(0)),
      formula(0, 0, integer(9), 9)
    ]
  );
  const xml = await z.get('xl/workbook.xml').text();
  assert.match(
    xml,
    /<calcPr fullPrecision="0" calcMode="manual" iterateCount="200" refMode="R1C1" iterateDelta="0.125" iterate="1" calcOnSave="0"\/>/
  );
});

test('empty NAME records retain both function calls and unresolved value references', () => {
  const named = cat([0x23], words(1, 0), integer(2), [0x22, 2], words(255));
  const { book, result, losses } = prepare(
    [defined('MyFunction', Buffer.alloc(0))],
    [formula(0, 0, named, 4), formula(0, 1, cat([0x23], words(1, 0)), 0)]
  );
  assert.equal(book.sheets[0].cells[0].formulaOut.text, 'MyFunction(2)');
  assert.equal(result.defined, '');
  assert.equal(book.sheets[0].cells[1].formulaOut.text, 'MyFunction');
  assert.equal(result.stats.converted, 2);
  assert.equal(result.stats.dropped, 0);
  assert.equal(losses.length, 0);
});

test('external name formulas use their own direct sheet-index pair and keep sheet scope', async () => {
  const ext = cat([0x3b], words(1, 1, 0, 2, 0, 1)),
    globals = [
      rec(0x1ae, cat(words(2), str('Other.xls'), str('First'), str('Second'))),
      rec(0x23, cat(words(0, 2, 0), [5, 0], Buffer.from('Costs'), words(ext.length), ext)),
      rec(0x17, cat(words(1), words(0, 1, 1)))
    ];
  const { z } = await converted(globals, [formula(0, 0, cat([0x39], words(0, 1, 0)), 7)]);
  assert.match(
    await z.get('xl/externalLinks/externalLink1.xml').text(),
    /<definedName name="Costs" refersTo="Second!\$A\$1:\$B\$3" sheetId="1"\/>/
  );
  assert.match(await z.get('xl/worksheets/sheet1.xml').text(), /<f>\[1\]Second!Costs<\/f>/);
});

test('Office rejects missing required function arguments; preserve only their cached results', () => {
  const call = cat(integer(2), [0x22, 1], words(345));
  const { book, result, losses } = prepare(
    [],
    [formula(0, 0, call, 29, undefined, 'e'), formula(0, 1, integer(9), 9)]
  );
  assert.equal(result.stats.dropped, 1);
  assert.equal(result.stats.converted, 1);
  assert.equal(book.sheets[0].cells[0].value, 29);
  assert.match(Object.keys(losses[0].detail.reasons)[0], /too few arguments/);
  assert.throws(() => X.formula(cat(integer(2), [0x22, 1], words(345))), { code: 'FORMULA_UNSUPPORTED' });
});

test('unbound registered functions do not silently become built-in functions of the same name', () => {
  const call = cat([0x23], words(1, 0), integer(2), [0x22, 2], words(255));
  const { book, result, losses } = prepare(
    [defined('Sum', Buffer.alloc(0))],
    [formula(0, 0, call, 29, undefined, 'e')]
  );
  assert.equal(result.stats.dropped, 1);
  assert.equal(book.sheets[0].cells[0].value, 29);
  assert.equal(losses[0].notify, true);
  assert.match(Object.keys(losses[0].detail.reasons)[0], /same name as an Excel function/);
});

test('Office external links omit sheetDataSet when the source has no cached sheets', async () => {
  const { z } = await converted(
    [rec(0x1ae, cat(words(1), str('Other.xls'), str('Data'))), rec(0x17, cat(words(1), words(0, 0, 0)))],
    [formula(0, 0, cat([0x3a], words(0, 0, 0)), 0)]
  );
  const xml = await z.get('xl/externalLinks/externalLink1.xml').text();
  assert.match(xml, /<sheetNames>/);
  assert.doesNotMatch(xml, /<sheetDataSet/);
});

test('Office plain workbooks omit unsupported Excel macro names with an explicit loss', () => {
  const macro = cat(integer(63), ref(0, 0), [0x22, 2], words(185));
  const { result, losses, book } = prepare(
    [defined('CellColour', macro)],
    [formula(0, 0, cat([0x23], words(1, 0)), 5)]
  );
  assert.equal(result.defined, '');
  assert.equal(result.stats.dropped, 1);
  assert.equal(book.sheets[0].cells[0].value, 5);
  assert.equal(losses[0].id, 'legacy-defined-name');
  assert.equal(losses[0].place, 'CellColour');
  assert.equal(losses[0].notify, true);
  assert.match(losses[0].detail.message, /macro-function/);
});

test('self-book NameX preserves a name scoped to another worksheet', () => {
  const globals = [
    defined('LocalOnly', integer(7), 0),
    rec(0x1ae, words(2, 0x401)),
    rec(0x17, cat(words(1), words(0, 0, 0)))
  ];
  const book = X.read(
      workbook(globals, [
        { name: 'One', records: [] },
        { name: 'Two', records: [formula(0, 0, cat([0x39], words(0, 1, 0)), 7)] }
      ]),
      { rich: true }
    ),
    losses = [];
  X.prepareFormulas(book, losses);
  assert.equal(book.sheets[1].cells[0].formulaOut.text, 'One!LocalOnly');
  assert.equal(losses.length, 0);
});

test('unbound ordinary names keep computed criteria formulas and their cached errors', () => {
  const cases = [
    ['glasses', cat([0x23], words(1, 0), [0x1d, 1, 0x0b]), 'glasses=TRUE'],
    ['age', cat([0x23], words(2, 0), integer(7), [0x09]), 'age<7'],
    ['size', cat([0x23], words(3, 0), integer(7), [0x06], integer(10), [0x0b]), 'size/7=10']
  ];
  const { book, result, losses } = prepare(
    cases.map(([name]) => defined(name, Buffer.alloc(0))),
    cases.map(([, tokens], col) => formula(46, col + 7, tokens, 29, undefined, 'e'))
  );
  assert.equal(result.stats.dropped, 0);
  assert.deepEqual(
    book.sheets[0].cells.map(c => c.formulaOut.text),
    cases.map(c => c[2])
  );
  assert.ok(book.sheets[0].cells.every(c => c.type === 'e' && c.value === 29));
  assert.equal(losses.length, 0);
});
