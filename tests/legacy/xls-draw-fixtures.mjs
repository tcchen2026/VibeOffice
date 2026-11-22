import { load } from '../../tools/legacy/load.mjs';
import { words, dword, rec, str, workbook, cfb } from './xls-fixtures.mjs';
const B = load(), X = B.xls;
const cat = (...items) => Buffer.concat(items.map(item => Buffer.from(item)));
const escher = (id, bytes, options = 0) => cat(words(options, id), dword(bytes.length), bytes);
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aXioAAAAASUVORK5CYII=', 'base64');
function properties(entries) {
  const heads = [], tails = [];
  for (const [id, value] of entries) {
    heads.push(words(id | (typeof value === 'string' ? 0x8000 : 0)));
    const bytes = typeof value === 'string' ? Buffer.from(value + '\0', 'utf16le') : null;
    heads.push(dword(bytes ? bytes.length : value));
    if (bytes) tails.push(bytes);
  }
  return escher(0xf00b, cat(...heads, ...tails), entries.length << 4 | 3);
}
function container(id, pieces) {
  const length = pieces.reduce((sum, part) => sum + (Buffer.isBuffer(part) ? part.length : 0), 0);
  return [cat(words(15, id), dword(length)), ...pieces];
}
function shape({ id = 1025, type = 1, objectType = 2, text, props = [], flags = 0xa00,
  anchor = words(0, 1, 256, 2, 128, 3, 512, 5, 64), child = false, coord, objectParts = [] } = {}) {
  const header = escher(0xf00a, cat(dword(id), dword(flags)), type << 4 | 2);
  const object = rec(0x5d, cat(words(0x15, 18, objectType, id & 1023), words(0x6011), Buffer.alloc(12),
    ...objectParts));
  const parts = [header, properties(props)];
  if (coord) parts.push(escher(0xf009, cat(...coord.map(dword))));
  parts.push(escher(child ? 0xf00f : 0xf010, anchor), escher(0xf011, []), { records: [object] });
  if (text != null) {
    const txo = Buffer.alloc(18);
    txo.writeUInt16LE(0x22, 0);
    txo.writeUInt16LE(text.length, 10);
    txo.writeUInt16LE(16, 12);
    parts.push(escher(0xf00d, []), { records: [rec(0x1b6, txo),
      rec(0x3c, cat([1], Buffer.from(text, 'utf16le'))),
      rec(0x3c, cat(words(0, 0), dword(0), words(text.length, 0), dword(0)))] });
  }
  return container(0xf004, parts);
}
function drawing(pieces) {
  const result = [];
  let bytes = [];
  for (const part of pieces) {
    if (Buffer.isBuffer(part)) bytes.push(part);
    else {
      if (bytes.length) result.push(rec(0xec, cat(...bytes)));
      bytes = [];
      result.push(...part.records);
    }
  }
  if (bytes.length) result.push(rec(0xec, cat(...bytes)));
  return result;
}
function pictureStore(defaults = []) {
  const image = escher(0xf01e, cat(Buffer.alloc(16), [255], png), 0x6e0 << 4);
  const header = Buffer.alloc(36);
  header.writeUInt32LE(image.length, 20);
  header.writeUInt32LE(1, 24);
  const store = escher(0xf001, escher(0xf007, cat(header, image), 2), 0x1f);
  return rec(0xeb, escher(0xf000, cat(store, properties(defaults)), 15));
}
async function convert(records, globals = []) {
  const output = await X.convert(cfb(workbook(globals, [{ name: 'Data', records }])));
  const zip = await L.zip.read(output.pkg);
  return { output, zip, xml: zip.has('xl/drawings/drawing1.xml') ?
    await zip.get('xl/drawings/drawing1.xml').text() : '' };
}

export { B, X, cat, escher, properties, shape, container, drawing, pictureStore, convert, png };
