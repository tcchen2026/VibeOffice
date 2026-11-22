import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import { load } from '../../tools/legacy/load.mjs';
import { encrypted } from './encrypted-fixtures.mjs';
import { doc } from './doc-fixtures.mjs';
import { cfb } from './ppt-fixtures.mjs';
const B = load(), E = B.encryption;
const vectors = JSON.parse(fs.readFileSync(new URL('./encryption-vectors.json', import.meta.url)));
const read = name => new Uint8Array(fs.readFileSync('/work/corpora/' + name));
const bytes = (value, encoding = 'hex') => new Uint8Array(Buffer.from(value, encoding));
const fixture = fragment => encrypted.find(([, name]) => name.includes(fragment));
const source = fragment => read(fixture(fragment)[1]);
const textParts = async pkg => {
  const zip = await globalThis.L.zip.read(pkg), parts = {};
  for (const name of zip.keys()) if (name.endsWith('.xml')) parts[name] = await zip.get(name).text();
  return parts;
};

test('legacy MD5 matches independent OpenSSL hashes across padding boundaries', () => {
  for (const length of [0, 1, 55, 56, 63, 64, 65, 119, 120, 512, 10000]) {
    const data = Uint8Array.from({ length }, (_, i) => (i * 37 + 11) & 255);
    assert.equal(Buffer.from(E.md5(data)).toString('hex'), createHash('md5').update(data).digest('hex'));
  }
});
test('legacy RC4 keeps stream position across calls: RFC 6229 40-bit key', () => {
  const cipher = E.rc4(bytes('0102030405'));
  assert.equal(Buffer.from(cipher(new Uint8Array(16))).toString('hex'), 'b2396305f03dc027ccc3524a0a1118a8');
  cipher(new Uint8Array(240));
  assert.equal(Buffer.from(cipher(new Uint8Array(16))).toString('hex'), '1cfcf62b03eddb641d77dfcf7f8d8c93');
});
for (const [i, vector] of vectors.headers.entries()) {
  test(`OpenSSL verifier ${i + 1}: ${vector.mode}, ${vector.bits} bits, Unicode/empty/long passwords`, () => {
    const info = E.header(bytes(vector.header));
    assert.equal(E.unlock(info, vector.password).mode, vector.mode);
    assert.throws(() => E.unlock(info, 'wrong'), { code: 'WRONG_PASSWORD' });
    assert.throws(() => E.unlock(info), { code: 'PASSWORD_REQUIRED' });
  });
}
for (const vector of vectors.workbooks) {
  test(`Excel ${vector.mode} preserves clear fields and rekeys across record bodies and padding`, async () => {
    const original = cfb({ Workbook: bytes(vector.bytes, 'base64') });
    const result = B.xls.decrypt(B.cfb(original), { password: vector.password });
    const expected = bytes(vector.plaintext, 'base64');
    assert.deepEqual(result.stream.subarray(0, expected.length), expected);
    assert.ok(result.stream.subarray(expected.length).every(v => v === 0));
    const converted = await B.xls.convert(original, { password: vector.password });
    const parts = await textParts(converted.pkg);
    assert.match(parts['xl/worksheets/sheet1.xml'], /<v>17<\/v>/);
    assert.match(parts['xl/worksheets/sheet1.xml'], /<v>42<\/v>/);
    assert.match(parts['xl/worksheets/sheet1.xml'], /<v>99<\/v>/);
  });
}
test('Word XOR uses its own rotation, preserves the FIB prefix and resets each stream', async () => {
  const original = doc({ text: 'XOR text\r', data: Buffer.from([0, 0x95, 5, 0x75, 0, 255]) });
  const plain = B.cfb(original), streams = {};
  // abc's independent FILEPASS values: key 514a, verifier cc1a. Word rotates the array by 7.
  const mask = bytes('95999475da577857da7465a87a2f2577');
  for (const name of ['WordDocument', '1Table', 'Data']) {
    const data = plain.stream(name), out = data.slice(), clear = name === 'WordDocument' ? 68 : 0;
    for (let i = clear; i < data.length; i++) {
      const xor = data[i] ^ mask[i & 15];
      if (data[i] && xor) out[i] = xor;
    }
    if (clear) {
      const v = new DataView(out.buffer);
      v.setUint16(10, v.getUint16(10, true) | 0x8100, true);
      v.setUint32(14, 0x514acc1a, true);
    }
    streams[name] = out;
  }
  const encrypted = cfb(streams), view = B.doc.decrypt(B.cfb(encrypted), { password: 'abc' });
  assert.deepEqual(view.stream('Data'), plain.stream('Data'));
  assert.equal(B.doc.read(view).storyText.main, 'XOR text\r');
  assert.match((await textParts((await B.doc.convert(encrypted, { password: 'abc' })).pkg))['word/document.xml'],
    /XOR text/);
  await assert.rejects(B.doc.convert(encrypted, { password: 'wrong' }), { code: 'WRONG_PASSWORD' });
});

for (const [format, name, password, expected] of encrypted) {
  test(`published encrypted ${format} fixture: ${name.split('__').at(-1)}`, async t => {
    if (!fs.existsSync('/work/corpora/' + name)) return t.skip('Pinned legacy corpus is not installed');
    const data = read(name), original = createHash('sha256').update(data).digest('hex');
    assert.equal(B.sniff(data).encrypted, true);
    await assert.rejects(B[format].convert(data, { password: 'I24 wrong password' }), { code: 'WRONG_PASSWORD' });
    if (password != null) await assert.rejects(B[format].convert(data), { code: 'PASSWORD_REQUIRED' });
    const result = await B[format].convert(data, { password });
    assert.equal(result.type, format === 'doc' ? 'docx' : format === 'xls' ? 'xlsx' : 'pptx');
    assert.equal(createHash('sha256').update(data).digest('hex'), original);
    const parts = await textParts(result.pkg);
    if (expected) assert.ok(Object.values(parts).some(part => part.includes(expected)), expected);
    if (name.includes('xor-encryption')) {
      assert.match(parts['xl/worksheets/sheet1.xml'], /<v>1<\/v>/);
      assert.match(parts['xl/worksheets/sheet1.xml'], /<v>2<\/v>/);
      assert.match(parts['xl/worksheets/sheet1.xml'], /<v>3<\/v>/);
    }
    if (name.includes('50833')) assert.match(parts['xl/comments1.xml'], /Robert Lawrence/);
    if (name.includes('png_encrypted')) assert.equal(result.stats.drawings.pictures, 1);
  });
}

test('PowerPoint encrypted picture bytes match all seven independent POI SHA-1 values', async () => {
  const view = B.ppt.decrypt(B.cfb(source('proc2356')), { password: 'crypto' });
  const pictures = new B.Reader(view.stream('Pictures'));
  const expected = [[0, 'nKsDTKqxTCR8LFkVVWlP9GSTvZ0='], [95163, 'SuNOR+9V1UVYZIoeD65l3VTaLoc='],
    [100864, 'Ql3IGrr4bNq07ZTp5iPg7b+pva8='], [714114, '8pdst9NjBGSfWezSZE8+aVhIRe0='],
    [723752, 'go6xqW7lvkCtlOO5tYLiMfb4oxw='], [770128, 'gZUM8YqRNL5kGNfyyYvEEernvCc='],
    [957958, 'CNU2iiqUFAnk3TDXsXV1ihH9eRM=']];
  for (const [offset, hash] of expected) {
    const picture = await B.officeart.blip(B.officeart.record(pictures, offset));
    assert.equal(createHash('sha1').update(picture.bytes).digest('base64'), hash);
  }
  const properties = B.properties(view.stream('\x05SummaryInformation'));
  assert.equal(properties[0].values[2].value, 'RC4 CryptoAPI Encryption');
});
test('damaged encrypted properties do not reject a presentation or notify a content loss', async () => {
  const raw = B.cfb(source('Protected-hello'));
  const broken = { ...raw, stream: name => name.toLowerCase() === 'encryptedsummary'
    ? new Uint8Array(4) : raw.stream(name) };
  const view = B.ppt.decrypt(broken, { password: 'hello' });
  const book = B.ppt.structure(view, []), result = await B.ppt.write(book, view, {});
  assert.equal(result.stats.slides, 1);
  assert.ok(result.losses.length > 0);
  assert.ok(result.losses.every(loss => loss.notify === false));
});
test('unused encrypted persist offsets are validated only when requested', () => {
  const view = B.ppt.decrypt(B.cfb(source('Protected-hello')), { password: 'hello' });
  view.pptDirectory.persists.set(900000, 0xffffffff);
  const book = B.ppt.read(view);
  assert.equal(book.slides.length, 1);
  assert.throws(() => book.get(900000), { name: 'LegacyError', code: 'TRUNCATED' });
});
test('CryptoAPI reserved values and provider names do not cause refusals; sizes and algorithms do', () => {
  const data = bytes(vectors.headers.find(v => v.mode === 'cryptoapi').header), view = new DataView(data.buffer);
  view.setUint32(12 + 4, 0xffffffff, true);
  view.setUint32(12 + 24, 0xffffffff, true);
  view.setUint32(12 + 28, 0xffffffff, true);
  assert.equal(E.unlock(E.header(data), 'Café漢字😀').mode, 'cryptoapi');
  view.setUint32(12 + 8, 0x660e, true);
  assert.throws(() => E.header(data), { code: 'ENCRYPTION_UNSUPPORTED' });
  view.setUint32(8, 0xffffffff, true);
  assert.throws(() => E.header(data), { name: 'LegacyError', code: 'TRUNCATED' });
  for (let n = 0; n < 52; n++) assert.throws(() => E.header(bytes(vectors.headers[0].header).subarray(0, n)),
    { name: 'LegacyError' });
});
test('BIFF5 default XOR decrypts before the separate unsupported macro-sheet refusal', async () => {
  const data = read('legacy-excel/poi__de9446dc6532__60284.xls');
  const decrypted = B.xls.decrypt(B.cfb(data));
  const book = B.xls.read(decrypted.stream, { decrypted: true });
  assert.equal(book.version, 'BIFF5/7');
  assert.ok(book.sheets.some(sheet => sheet.name === 'Module1'));
  await assert.rejects(B.xls.convert(data), { code: 'SHEET_KIND' });
});
