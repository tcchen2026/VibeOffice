// Authored OLE/SmartArt objects: foreign clipboard, content edits, undo and drafts.
// node tools/quire/test/objects.mjs ~/corpora/word /tmp/quire-object-copies
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [corpus, output] = process.argv.slice(2);
if (!corpus || !output) throw new Error('Usage: objects.mjs CORPUS OUTPUT');
const cases = [
  ['libreoffice__c625088ae100__test_ole_object.docx', 'OLEObject'],
  ['libreoffice__32d19c94369d__smartart.docx', 'relIds'],
];
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), source = await openPage('quire', { persistence: false }), reports = [];
let attempted;
try {
  for (const [file, feature] of cases) {
    attempted = file;
    const data = fs.readFileSync(path.join(corpus, file)).toString('base64');
    const clipboard = await source.evaluate(`(async () => {
      const d = (await L.docx.read(Uint8Array.from(atob(${JSON.stringify(data)}), c => c.charCodeAt(0)))).doc;
      L.app.loadDoc(d, 'Source', {saved:true});
      const blocks = L.D.cloneBlocks(d.main.blocks);
      return { blocks, preserved: L.preserve.clipboard(blocks) };
    })()`);
    const destination = await openPage('quire', { persistence: false });
    try {
      const result = await destination.evaluate(`(${check.toString()})(${JSON.stringify(clipboard)}, ${JSON.stringify(feature)})`);
      for (const [scenario, bytes] of Object.entries(result.artifacts)) {
        fs.mkdirSync(path.join(output, scenario), { recursive: true });
        fs.writeFileSync(path.join(output, scenario, file), Buffer.from(bytes, 'base64'));
        reports.push({ file, scenario, attempted: true, status: 'ok', checks: result.checks });
      }
      if (destination.errors.length) throw new Error(destination.errors.join('\n'));
      console.log(file + ': ' + result.checks + ' checks');
    } finally { await destination.close(); }
  }
  if (source.errors.length) throw new Error(source.errors.join('\n'));
} catch (error) {
  reports.push({ file: attempted, scenario: 'clipboard', attempted: true, status: 'failed', error: error.message });
  throw error;
} finally {
  fs.writeFileSync(path.join(output, 'results.jsonl'), reports.map(r => JSON.stringify(r)).join('\n') + '\n');
  await source.close(); stop();
}

async function check(clipboard, feature) {
  const { D, O, preserve: P, opc: K } = L, artifacts = {};
  let checks = 0;
  const assert = (ok, message) => { if (!ok) throw new Error(message); checks++; };
  const encode = async blob => { const b = new Uint8Array(await blob.arrayBuffer()); let s = ''; for (let i = 0; i < b.length; i += 32768) s += String.fromCharCode(...b.subarray(i, i + 32768)); return btoa(s); };
  const read = async blob => (await L.docx.read(new Uint8Array(await blob.arrayBuffer()))).doc;
  const xml = d => K.raw(d.pkg.xml(d.pkg.main));
  const objects = d => { const out = []; D.walk(d.main, p => { if (p.t === 'p') for (const it of p.runs) if (it.keep?.opaque && ['group', 'shape', 'img'].includes(it.t)) out.push({ p, it }); }); return out; };
  const save = async (kind, doc, blob) => { blob ||= await L.docx.write(doc); artifacts[kind] = await encode(blob); return read(blob); };
  const doc = D.newDoc(); L.app.loadDoc(doc, 'Pasted objects', { saved: true });
  const blocks = await P.hydrateClipboard(K.remapSources(clipboard.blocks, K.import(clipboard.preserved)));
  D.tx('Paste objects', () => O.insertBlocks(D.pos(doc.main.blocks[0], 0), blocks));
  const list = objects(doc); assert(list.length, 'Clipboard lost the model objects');
  const images = list.filter(o => o.it.media);
  assert(images.every(o => L.media.has(o.it.media)), 'Foreign clipboard lost the preview bytes');
  const copied = await save('copied', doc);
  assert(xml(copied).includes(':' + feature), 'Foreign clipboard converted the object');
  assert(!doc.losses?.length, 'Foreign clipboard reported an unexpected conversion');
  if (feature === 'relIds') {
    const rels = copied.pkg.rels(copied.pkg.main), data = rels.filter(r => /\/diagramData$/.test(r.type));
    assert(data.length && data.every(r => Array.from(copied.pkg.xml(r.part).getElementsByTagName('*')).filter(e => e.localName === 'dataModelExt').every(e => rels.some(link => link.id === e.getAttribute('relId') && /\/diagramDrawing$/.test(link.type)))), 'Copied SmartArt has a broken implicit drawing link');
  }
  const { p, it } = list[0];
  D.tx('Edit converted content', () => { D.touch(p); if (it.t === 'group') it.kids[0].fill = { t: 'solid', c: '#AB1234' }; else it.media = null; });
  await save('converted', doc);
  assert(doc.losses?.some(l => l.action === 'conversion'), 'Content replacement was not reported');
  assert(D.undo(), 'Missing undo transaction');
  const undone = await save('undo', doc);
  assert(xml(undone).includes(':' + feature), 'Undo failed to restore the object');
  assert(!doc.losses?.length, 'Undo left a stale writer loss');
  const draft = await save('draft', doc, await window.__corpusHooks.snapshot(doc));
  const recovered = await save('recovered', draft);
  assert(xml(recovered).includes(':' + feature), 'Draft recovery converted the object');
  return { checks, artifacts };
}
