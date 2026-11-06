import { deflateSync } from 'node:zlib';
import { doc } from './doc-fixtures.mjs';
import { workbook, rec as biff, str } from './xls-fixtures.mjs';
import { presentation, container, rec, u16, u32 } from './ppt-fixtures.mjs';

const cat = (...parts) => Buffer.concat(parts);
export function source(streams, metadata = {}) {
  const entries = new Map();
  for (const [path, data] of Object.entries(streams)) {
    const parts = path.split('/');
    for (let i = 1; i < parts.length; i++) {
      const parent = parts.slice(0, i).join('/');
      entries.set(parent, { path: parent, name: parts[i - 1], type: 1, size: 0, ...metadata[parent] });
    }
    entries.set(path, { path, name: parts.at(-1), type: 2, size: data.length, ...metadata[path] });
  }
  return { entries: [...entries.values()], root: { name: 'Root Entry', type: 5, path: '' },
    stream: path => streams[path] ?? null };
}

// Literal-only compressed data is an independent MS-OVBA fixture. Chunk boundaries
// and flag bytes are explicit; it never uses the production decompressor to encode.
export function compressed(bytes) {
  const payload = [];
  for (let p = 0; p < bytes.length; p += 8) payload.push(Buffer.from([0]), bytes.subarray(p, p + 8));
  const data = cat(...payload);
  if (data.length > 4096) throw new Error('Fixture chunk too large');
  return cat(Buffer.from([1]), u16(0xb000 | (data.length - 1)), data);
}
export function project(overrides = {}) {
  return {
    PROJECT: Buffer.from('ID="{00000000-0000-0000-0000-000000000001}"\r\n' +
      'Document=ThisWorkbook/&H00000000\r\nDocument=SheetCode/&H00000000\r\n' +
      'Module=Module1\r\nName="Project"\r\n'),
    'VBA/dir': compressed(cat(u16(3), u32(2), u16(1252), u16(0x10), u32(0))),
    'VBA/_VBA_PROJECT': Buffer.from([0xcc, 0x61, 0xff, 0xff, 0, 0, 0]),
    'VBA/Module1': Buffer.from('An opaque module fixture; it is never run.'),
    ...overrides
  };
}
const rooted = (streams, prefix) => Object.fromEntries(Object.entries(streams).map(([p, b]) => [prefix + p, b]));
const copy = cfb => Object.fromEntries(cfb.entries.filter(e => e.type === 2).map(e => [e.path, cfb.stream(e.path)]));
export function word(B, pack, options = {}, streams = project()) {
  return pack(source({ ...copy(B.cfb(doc(options))), ...rooted(streams, 'Macros/') }));
}
export function excel(pack, options = {}, streams = project()) {
  const globals = [biff(0x1ba, str(options.workbookName || 'ThisWorkbook'))];
  if (options.template) globals.push(biff(0x60));
  const data = workbook(globals, [{ name: 'Visible title', records: [biff(0x1ba, str(options.sheetName || 'SheetCode'))] }]);
  return pack(source({ Workbook: data, ...rooted(streams, '_VBA_PROJECT_CUR/') }));
}
export function powerpoint(pack, options = {}, streams = project()) {
  const bytes = pack(source(streams));
  const payload = options.uncompressed ? bytes : cat(u32(bytes.length), deflateSync(bytes));
  const atom = rec(1024, cat(u32(options.reference || 40), u32(options.empty ? 0 : 1), u32(2)), 0, 2);
  return presentation({ pages: options.pages, documentRecords: [container(2000, [container(1023, [atom], 1)])],
    extra: [{ ref: 40, record: rec(4113, payload, options.uncompressed ? 0 : 1) }] }).bytes;
}
export function commands(name = 'Project.Module1.Run') {
  const descriptor = cat(Buffer.from([0x56, 0]), u16(7), u16(0), u16(0xffff), Buffer.alloc(16));
  const strings = cat(u16(0xffff), u16(1), u16(2), u16(name.length), Buffer.from(name, 'utf16le'), u16(1));
  const upper = name.toUpperCase();
  const names = cat(u16(1), u16(7), u16(upper.length), Buffer.from(upper, 'utf16le'), u16(0));
  return cat(Buffer.from([255, 1]), u32(1), descriptor, Buffer.from([16]), strings,
    Buffer.from([17]), names, Buffer.from([64]));
}
