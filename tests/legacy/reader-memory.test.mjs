import { test } from 'node:test';
import assert from 'node:assert/strict';
import { load } from '../../tools/legacy/load.mjs';
import { rec, words, workbook, cfb } from './xls-fixtures.mjs';
const B = load();

test('record windows keep local bounds with a bounded backing-view allocation count', () => {
  const NativeView = globalThis.DataView;
  let allocated = 0;
  globalThis.DataView = class extends NativeView {
    constructor(...args) {
      super(...args);
      allocated++;
    }
  };
  try {
    const input = Uint8Array.from({ length: 80004 }, (_, i) => i & 255);
    const root = new B.Reader(input.subarray(4), 'records');
    for (let i = 0; i < 20000; i++) {
      const record = root.reader(i * 4, 4),
        nested = record.reader(1, 2);
      assert.equal(nested.u16(0), input[i * 4 + 5] | (input[i * 4 + 6] << 8));
      assert.deepEqual([...nested.bytes], [...input.subarray(i * 4 + 5, i * 4 + 7)]);
      assert.throws(() => nested.u16(1), { name: 'LegacyError', code: 'TRUNCATED' });
    }
    assert.ok(allocated <= 2, `${allocated} backing views for 20,000 record windows`);
  } finally {
    globalThis.DataView = NativeView;
  }
});

test('workbook identification reads globals without parsing a damaged worksheet body', async () => {
  const stream = workbook([], [{ name: 'Data', records: [rec(0x203, words(0, 0, 0))] }]);
  // Append a record whose declared body is beyond the stream, after the globals.
  const bytes = cfb(Buffer.concat([stream, Buffer.from([0x03, 0x02, 0xff, 0xff])]));
  assert.deepEqual(B.sniff(bytes), { format: 'xls', version: 'BIFF8', encrypted: false });
  await assert.rejects(B.xls.convert(bytes), { name: 'LegacyError', code: 'TRUNCATED' });
});

test('ordered and unsorted cell streams keep the same last value for repeated addresses', async () => {
  const number = (row, col, value) => {
    const bytes = Buffer.alloc(8);
    bytes.writeDoubleLE(value);
    return rec(0x203, Buffer.concat([words(row, col, 0), bytes]));
  };
  const sheet = async records => {
    const result = await B.xls.convert(cfb(workbook([], [{ name: 'Data', records }])));
    return await (await L.zip.read(result.pkg)).get('xl/worksheets/sheet1.xml').text();
  };
  const ordered = await sheet([number(0, 0, 1), number(0, 1, 2), number(2, 0, 3)]);
  const unsorted = await sheet([number(2, 0, 99), number(0, 1, 2), number(0, 0, 1), number(2, 0, 3)]);
  assert.equal(unsorted, ordered);
  assert.deepEqual(
    [...ordered.matchAll(/<c r="([^"]+)"><v>([^<]+)<\/v><\/c>/g)].map(([, address, value]) => [
      address,
      value
    ]),
    [
      ['A1', '1'],
      ['B1', '2'],
      ['A3', '3']
    ]
  );
});
