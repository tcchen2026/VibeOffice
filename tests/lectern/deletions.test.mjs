// What the user deletes takes its parts along, recorded as deleted (not as the writer's loss alarm, unreferenced:);
// an unedited save records nothing of the kind.
// node --test tests/lectern/deletions.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const corpus = path.join(process.env.VO_CORPORA || '/work/corpora', 'powerpoint');
const slides = path.join(corpus, 'openxml-sdk__88d7f7c81442__typical.pptx'); // slide 2 has notes
const pictures = path.join(corpus, 'libreoffice__18fd956ba0f3__tdf164622.pptx'); // the first shape is an SVG picture with a PNG fallback
let page, stop;
before(async () => {
  if (!fs.existsSync(slides) || !fs.existsSync(pictures)) return;
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

async function save(data, kind) {
  const pres = await L.pptx.read(Uint8Array.from(atob(data), c => c.charCodeAt(0))); L.pres = pres;
  if (kind === 'slide') pres.slides.splice(1, 1);
  if (kind === 'picture') pres.slides[0].shapes.shift();
  if (kind === 'shared') { const list = pres.slides[0].shapes; list.push(L.model.dup(list[0])); list.shift(); } // its copy still shows the picture
  await L.pptx.write(pres, { format: 'pptx' });
  const ids = prefix => (pres.losses || []).filter(e => e.id.startsWith(prefix)).map(e => e.where).sort();
  return { deleted: ids('deleted:'), alarms: ids('unreferenced:'), shown: (pres.losses || []).filter(e => e.notify && /^deleted:/.test(e.id)).length };
}
const run = async (file, kind) => {
  const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(fs.readFileSync(file).toString('base64'))}, ${JSON.stringify(kind)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

test('a deleted slide takes its notes; an unedited save records no deletion', async t => {
  if (!page) return t.skip('Fetch pinned corpus: ' + slides);
  assert.deepEqual(await run(slides, 'unedited'), { deleted: [], alarms: [], shown: 0 });
  const result = await run(slides, 'slide');
  assert.deepEqual(result.alarms, []);
  assert.ok(result.deleted.includes('ppt/slides/slide2.xml'), result.deleted.join(' '));
  assert.equal(result.shown, 0, 'the user deleted it: nothing to tell');
});

test('a deleted picture takes its files; a picture still used elsewhere is not explained away', async t => {
  if (!page) return t.skip('Fetch pinned corpus: ' + pictures);
  assert.deepEqual(await run(pictures, 'unedited'), { deleted: [], alarms: [], shown: 0 });
  const result = await run(pictures, 'picture');
  assert.deepEqual(result.alarms, []);
  assert.deepEqual(result.deleted, ['ppt/media/image1.png', 'ppt/media/image2.svg']);
  const shared = await run(pictures, 'shared');
  assert.deepEqual(shared, { deleted: [], alarms: [], shown: 0 });
});
