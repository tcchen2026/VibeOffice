// Isolate identity reservation on a large worksheet containing ordinary revision words.
// node --expose-gc tools/ooxml/prescan-bench.mjs [CELLS=100000] [OPC_SOURCE]
import { performance } from 'node:perf_hooks';
import { pathToFileURL } from 'node:url';
import path from 'node:path';
import '../../public/common/xml.js';
import '../../public/common/zip.js';
const count = +(process.argv[2] || 100000), source = process.argv[3];
if (!Number.isInteger(count) || count < 1 || count > 1048576) throw new Error('Invalid cell count');
await import(source ? pathToFileURL(path.resolve(source)).href : '../../public/common/opc.js');
await import('../../public/common/opc-order.js');
const K = L.opc, s = K.NS.s;
const worksheet = `<worksheet xmlns="${s}"><sheetData>` + Array.from({ length: count }, (_, i) =>
  `<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr"><is><t>del documento</t></is></c></row>`).join('') + '</sheetData></worksheet>';
const blob = await L.zip.write([
  { name: '[Content_Types].xml', data: `<Types xmlns="${K.NS.ct}"><Default Extension="xml" ContentType="application/xml"/></Types>` },
  { name: 'xl/worksheets/sheet1.xml', data: worksheet },
]);
const pkg = await K.open(new Uint8Array(await blob.arrayBuffer()));
// Warm byte decoding before measuring the identity pre-scan itself.
for (const name of pkg.names) pkg.text(name);
globalThis.gc?.();
const heap = process.memoryUsage().heapUsed, start = performance.now();
new K.Writer(pkg);
console.log(JSON.stringify({ cells: count, xmlBytes: worksheet.length, ms: performance.now() - start,
  heapGrowthMB: (process.memoryUsage().heapUsed - heap) / 1048576, source: source || 'current' }));
