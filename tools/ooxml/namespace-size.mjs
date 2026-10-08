// Isolate namespace compaction on existing saves; this is not an app round trip.
// node tools/ooxml/namespace-size.mjs ORIGINALS SAVED OUTPUT [LIMIT=400]
import fs from 'node:fs';
import path from 'node:path';
import '../../public/common/xml.js';
import '../../public/common/opc.js';
import '../../public/common/zip.js';
const [originals, saved, output, count = '400'] = process.argv.slice(2);
if (!output) throw new Error('Usage: namespace-size.mjs ORIGINALS SAVED OUTPUT [LIMIT=400]');
if (fs.existsSync(output) && fs.readdirSync(output).length) throw new Error('Use an empty output directory');
fs.mkdirSync(output, { recursive: true });
const files = fs.readdirSync(saved).filter(n => /\.doc[mx]$/i.test(n)).sort().slice(0, +count), rows = [];
const K = L.opc;
const main = async zip => K.parse(await zip.get('_rels/.rels').text()).children.find(r => /\/officeDocument$/.test(r.getAttribute('Type')))?.getAttribute('Target').replace(/^\//, '');
for (const file of files) {
  const row = { file, attempted: true };
  try {
    const a = await L.zip.read(fs.readFileSync(path.join(originals, file))), b = await L.zip.read(fs.readFileSync(path.join(saved, file)));
    const part = await main(b), xml = await b.get(part).text(), compact = K.hoistNamespaces(xml);
    row.original = (await a.get(await main(a)).bytes()).length;
    row.before = new TextEncoder().encode(xml).length; row.after = new TextEncoder().encode(compact).length;
    row.innerIgnorable = K.parse(compact).getElementsByTagName('*').filter(e => e.parentNode?.nodeType === 1 && e.getAttributeNS(K.NS.mc, 'Ignorable') != null).length;
    // Follow the writer-owned roles to compact the other rewritten Word parts
    // too; theme, glossary, custom XML and embedded graphs remain opaque.
    const rewritten = new Set([part]);
    const owned = new Set(['styles', 'numbering', 'settings', 'fontTable', 'footnotes', 'endnotes', 'comments', 'commentsExtended', 'header', 'footer', 'core-properties', 'extended-properties', 'custom-properties']);
    for (const [name, entry] of b) if (name.endsWith('.rels')) {
      for (const r of K.parse(await entry.text()).children) if (r.getAttribute('TargetMode') !== 'External' && owned.has(r.getAttribute('Type')?.split('/').pop()))
        rewritten.add(K.resolve(K.owner(name), r.getAttribute('Target')).part);
    }
    const entries = [];
    for (const [name, entry] of b) entries.push({ name, data: name === part ? compact : rewritten.has(name) ? K.hoistNamespaces(await entry.text()) : await entry.bytes() });
    row.rewritten = [...rewritten];
    const blob = await L.zip.write(entries); fs.writeFileSync(path.join(output, file), new Uint8Array(await blob.arrayBuffer()));
    row.status = 'ok'; row.part = part;
  } catch (e) { row.status = 'failed'; row.error = e.message; }
  rows.push(row);
}
fs.writeFileSync(path.join(output, 'results.jsonl'), rows.map(r => JSON.stringify(r)).join('\n') + '\n');
const ok = rows.filter(r => r.status === 'ok');
const summary = { attempted: rows.length, ok: ok.length, failed: rows.length - ok.length,
  originalBytes: ok.reduce((n, r) => n + r.original, 0), beforeBytes: ok.reduce((n, r) => n + r.before, 0), afterBytes: ok.reduce((n, r) => n + r.after, 0),
  innerIgnorable: ok.reduce((n, r) => n + r.innerIgnorable, 0) };
fs.writeFileSync(path.join(output, 'summary.json'), JSON.stringify(summary, null, 2) + '\n');
console.log(JSON.stringify(summary));
if (summary.failed || summary.innerIgnorable) process.exitCode = 1;
