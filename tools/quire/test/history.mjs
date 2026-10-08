// Real DOCX regression for typed chart bytes surviving text edits, undo and copying.
// Usage: node tools/quire/test/history.mjs FILE.docx
import fs from 'node:fs';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';

const file = process.argv[2];
if (!file) throw new Error('Usage: history.mjs FILE.docx (use the corpus chart-size.docx)');
const data = fs.readFileSync(file).toString('base64');
const stop = await ensureServer();
let page;
try {
  page = await openPage('quire');
  const result = await page.evaluate(`(async () => {
    const { D, O } = L;
    const { doc } = await L.docx.read(Uint8Array.from(atob(${JSON.stringify(data)}), c => c.charCodeAt(0)));
    D.doc = doc; D.resetHistory(); D.reindex(doc);
    let paragraph;
    D.walk(doc.main, p => { if (!paragraph && p.t === 'p' && p.runs.some(r => r.t === 'chart')) paragraph = p; });
    if (!paragraph) throw new Error('Fixture has no chart paragraph');
    const chart = paragraph.runs.find(r => r.t === 'chart');
    const source = typeof chart.src === 'string' ? L.chart.src.get(chart.src) : chart.src;
    if (!source?.files?.length) throw new Error('Fixture has no carried chart parts');
    const check = async () => {
      const blob = await L.docx.write(doc);
      const zip = await L.zip.read(new Uint8Array(await blob.arrayBuffer()));
      for (const f of source.files) {
        const part = f.path === source.path ? 'word/charts/chart1.xml' : f.path === L.dml.relsPath(source.path) ? 'word/charts/_rels/chart1.xml.rels' : f.path;
        const entry = zip.get(part);
        if (!entry) throw new Error('Missing ' + part);
        const bytes = await entry.bytes();
        if (bytes.length !== f.data.length || bytes.some((b, i) => b !== f.data[i])) throw new Error('Changed original chart bytes: ' + part);
      }
    };
    await check();
    D.tx('Type alongside chart', () => O.insertText(D.pos(paragraph, D.plen(paragraph)), ' HISTORY-CHECK'));
    await check();
    if (!D.undo()) throw new Error('Undo failed');
    await check();
    if (!D.redo()) throw new Error('Redo failed');
    await check();
    const snapshot = D.snapshot(doc);
    const copy = D.cloneBlocks([paragraph]);
    doc.main.blocks = copy;
    await check();
    D.restoreSnapshot(doc, snapshot);
    await check();
    return { parts: source.files.length, states: 6 };
  })()`);
  assert.deepEqual(page.errors, [], 'Browser errors');
  console.log(JSON.stringify({ file, status: 'ok', ...result }));
} finally { await page?.close(); stop(); }
