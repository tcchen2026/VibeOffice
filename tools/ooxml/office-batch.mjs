// Build a bounded manual Office review from authored inputs and explicit edits.
// BASE can point at an isolated previous writer for the required before/after diff.
// node tools/ooxml/office-batch.mjs MANIFEST.json OUTPUT_DIR
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { ensureServer, openPage } from './cdp.mjs';
const [manifest, output] = process.argv.slice(2);
if (!manifest || !output) throw new Error('Usage: office-batch.mjs MANIFEST.json OUTPUT_DIR');
const cases = JSON.parse(fs.readFileSync(manifest, 'utf8'));
if (!cases.length || cases.length > 15) throw new Error('An Office batch has 1–15 files');
if (fs.existsSync(output) && fs.readdirSync(output).length) throw new Error('Use an empty output directory');
fs.mkdirSync(output, { recursive: true });
const helper = fs.readFileSync(new URL('./scenarios.js', import.meta.url), 'utf8');
const stop = await ensureServer(), reports = [], pages = new Map();
try {
  for (const c of cases) {
    const row = { ...c, attempted: true, officeAcceptance: 'pending' };
    try {
      if (!pages.has(c.app)) { const page = await openPage(c.app, { persistence: false }); await page.evaluate(helper); pages.set(c.app, page); }
      const page = pages.get(c.app); page.errors.length = 0;
      const bytes = fs.readFileSync(c.source); row.sourceHash = createHash('sha256').update(bytes).digest('hex');
      const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(c)}, ${JSON.stringify(bytes.toString('base64'))})`);
      if (page.errors.length) throw new Error(page.errors.join('\n'));
      const saved = Buffer.from(result.data, 'base64'); delete result.data;
      fs.writeFileSync(path.join(output, c.file), saved);
      Object.assign(row, result, { status: 'ok', savedHash: createHash('sha256').update(saved).digest('hex') });
      console.log(c.file + ': saved');
    } catch (error) { Object.assign(row, { status: 'failed', error: error.stack }); console.error(c.file + ': ' + error.message); process.exitCode = 1; }
    reports.push(row);
    fs.writeFileSync(path.join(output, 'manifest.json'), JSON.stringify(reports, null, 2) + '\n');
  }
  const lines = ['# Office batch', '', 'Open each file in its Office application. Record ✓, **repair prompt**, or **looks wrong** and any details. For macro files, review project presence without running the code. Original links point to the authored inputs outside this small batch.', '', '| File | Feature / edit | What to look for | Result |', '|---|---|---|---|'];
  for (const c of reports) lines.push(`| [${c.file}](${c.file}) · [original](${c.source.replaceAll(' ', '%20')}) | ${c.feature} | ${c.lookFor} | ${c.status === 'ok' ? 'Pending' : 'Generation failed: ' + c.error.split('\n')[0]} |`);
  fs.writeFileSync(path.join(output, 'CHECKLIST.md'), lines.join('\n') + '\n');
} finally { for (const page of pages.values()) await page.close(); stop(); }

async function save(c, data) {
  const { opc: K } = L, bytes = Uint8Array.from(atob(data), x => x.charCodeAt(0));
  const read = async bytes => c.app === 'quire' ? (await L.docx.read(bytes)).doc : c.app === 'ledger' ? L.xlsxRead.read(bytes) : L.pptx.read(bytes);
  let doc = await read(bytes), changed = [];
  if (c.app === 'quire') {
    const { D, O, preserve: P } = L;
    L.app.loadDoc(doc, 'Office review', { saved: true });
    if (c.edit === 'binding') {
      const record = [...P.controls(doc).records.values()].find(r => r.complete && r.control.binding && r.control.type === 'text');
      if (!record) throw new Error('No editable bound text control');
      D.tx('Edit bound value', () => O.insertText(D.pos(record.start.pos.p, record.start.pos.o + 1), ' EDITED'));
      changed.push('Inserted EDITED into the bound text');
    } else if (c.edit === 'paste') {
      const blocks = D.cloneBlocks(doc.main.blocks), imported = await P.hydrateClipboard(K.remapSources(blocks, K.import(P.clipboard(blocks))));
      doc = D.newDoc(); L.app.loadDoc(doc, 'Pasted object', { saved: true });
      D.tx('Paste object', () => O.insertBlocks(D.pos(doc.main.blocks[0], 0), imported));
      changed.push('Pasted all source blocks into a new document');
    } else if (['picture-properties', 'shape-shadow', 'drawing-anchor'].includes(c.edit)) {
      let owner, object;
      D.walk(doc.main, p => { if (p.t === 'p' && !object) { const it = p.runs.find(it => c.edit === 'shape-shadow' ? it.t === 'shape' : it.t === 'img'); if (it) { owner = p; object = it; } } });
      if (!object) throw new Error('No drawing for ' + c.edit);
      D.tx('Edit drawing properties', () => {
        D.touch(owner);
        if (c.edit === 'picture-properties') { object.crop = { l: 0.15, t: 0.05, r: 0, b: 0 }; object.border = { val: 'single', sz: 1.5, color: '123456' }; object.alt = 'Edited picture'; }
        else if (c.edit === 'shape-shadow') object.shadow = { c: '#123456', a: 0.6, dx: 5, dy: 3, blur: 2 };
        else object.float = { ...(object.float || {}), wrap: 'square', posH: { rel: 'column', off: 12 }, posV: { rel: 'paragraph', off: 10 }, allowOverlap: false };
      });
      changed.push(c.edit === 'picture-properties' ? 'Cropped the picture and changed its border and alternative text' : c.edit === 'shape-shadow' ? 'Changed the outer shadow' : 'Changed the drawing anchor and wrapping');
    } else if (c.edit === 'watermark') {
      D.tx('Replace watermark', () => { D.touchKey(doc, 'watermark'); doc.watermark = { type: 'text', text: 'REVISED', color: 'C0C0C0', layout: 'diagonal' }; });
      changed.push('Replaced the watermark with REVISED');
    } else if (c.edit === 'wordart') {
      let owner, shape;
      D.walk(doc.main, p => { if (p.t === 'p' && !shape) { const s = p.runs.find(r => r.wordart); if (s) { owner = p; shape = s; } } });
      if (!shape) throw new Error('No editable WordArt');
      D.tx('Edit WordArt', () => { D.touch(owner); shape.wordart.text += ' EDITED'; });
      changed.push('Appended EDITED to WordArt');
    } else if (c.edit === 'revision-link') {
      let owner, run;
      D.walk(doc.main, p => { if (p.t === 'p' && !run) { const r = p.runs.find(r => r.t === 'text' && (r.rPr.ins || r.rPr.del)); if (r) { owner = p; run = r; } } });
      if (!run) throw new Error('No tracked text');
      D.tx('Link tracked text', () => { D.touch(owner); run.rPr.link = { url: 'https://example.invalid/review' }; });
      changed.push('Added a hyperlink to tracked text');
    } else if (c.edit && c.edit !== 'save') changed = corpusScenario(c.app, doc, c.edit).changed;
  } else if (c.edit && c.edit !== 'save') changed = corpusScenario(c.app, doc, c.edit).changed;
  const format = c.file.split('.').pop();
  const blob = await (c.app === 'quire' ? L.docx.write(doc, { format }) : c.app === 'ledger' ? L.xlsxWrite.write(doc, { type: format }) : L.pptx.write(doc, { format }));
  const saved = new Uint8Array(await blob.arrayBuffer());
  await read(saved);
  let encoded = ''; for (let i = 0; i < saved.length; i += 32768) encoded += String.fromCharCode(...saved.subarray(i, i + 32768));
  return { data: btoa(encoded), changed, losses: doc.losses || [] };
}
