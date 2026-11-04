// What the user deletes (a picture, a chart) is recorded as deleted, not as the writer's loss alarm (unreferenced:).
// node --test tests/quire/deletions.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const corpus = path.join(process.env.VO_CORPORA || '/work/corpora', 'word');
const picture = path.join(corpus, 'libreoffice__34e27588d443__fdo76597.docx');
const chart = path.join(corpus, 'libreoffice__150741726c0c__tdf155685.docx');
let page, stop;
before(async () => {
  if (!fs.existsSync(picture) || !fs.existsSync(chart)) return;
  stop = await ensureServer(process.env.BASE);
  page = await openPage('quire', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

async function save(data, kind) {
  const { D } = L, { doc } = await L.docx.read(Uint8Array.from(atob(data), c => c.charCodeAt(0)));
  D.doc = doc; D.resetHistory();
  let p; D.walk(doc.main, b => { if (!p && b.t === 'p' && b.runs.some(r => ['img', 'chart'].includes(r.t))) p = b; });
  const object = p.runs.find(r => ['img', 'chart'].includes(r.t));
  if (kind === 'deleted') D.tx('Delete', () => { D.touch(p); p.runs.splice(p.runs.indexOf(object), 1); });
  if (kind === 'moved') D.tx('Move', () => { D.touch(p); p.runs.splice(p.runs.indexOf(object), 1); doc.main.blocks.push(D.para([object])); });
  await L.docx.write(doc);
  const where = prefix => (doc.losses || []).filter(e => e.id.startsWith(prefix)).map(e => e.where).sort();
  return { deleted: where('deleted:'), alarms: where('unreferenced:') };
}
const run = async (file, kind) => {
  const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(fs.readFileSync(file).toString('base64'))}, ${JSON.stringify(kind)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

test('a deleted picture takes its file; a moved one, still used, records nothing', async t => {
  if (!page) return t.skip('Fetch pinned corpus: ' + picture);
  assert.deepEqual(await run(picture, 'unedited'), { deleted: [], alarms: [] });
  assert.deepEqual(await run(picture, 'deleted'), { deleted: ['word/media/image1.jpeg'], alarms: [] });
  assert.deepEqual(await run(picture, 'moved'), { deleted: [], alarms: [] });
});

test('a deleted chart takes its part and what it references', async t => {
  if (!page) return t.skip('Fetch pinned corpus: ' + chart);
  assert.deepEqual(await run(chart, 'unedited'), { deleted: [], alarms: [] });
  const result = await run(chart, 'deleted');
  assert.deepEqual(result.alarms, []);
  assert.deepEqual(result.deleted, ['word/charts/chart1.xml']);
});
