// A PowerPoint 97-2003 diagram (an organization chart in the slide's VML, com:legacyDrawing) is kept as it was.
// node --test tests/lectern/legacy-diagram.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const file = path.join(process.env.VO_CORPORA || '/work/corpora', 'powerpoint', 'poi__6d3225f18c39__PPTWithAttachments.pptm');
let page, stop;
before(async () => {
  if (!fs.existsSync(file)) return;
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

async function save(data, kind) {
  const pres = await L.pptx.read(Uint8Array.from(atob(data), c => c.charCodeAt(0))); L.pres = pres;
  const list = pres.slides[0].shapes, i = list.findIndex(s => s.name === 'Organization Chart 8'), shape = list[i];
  if (kind === 'moved') shape.x += 20;
  if (kind === 'restyled') shape.line = { t: 'solid', c: '#FF0000', w: 2 };
  if (kind === 'deleted') list.splice(i, 1);
  const pkg = await L.opc.open(new Uint8Array(await (await L.pptx.write(pres, { format: 'pptm' })).arrayBuffer()));
  const slide = pkg.text('ppt/slides/slide1.xml');
  const frame = (slide.match(/<p:graphicFrame>(?:(?!<\/p:graphicFrame>).)*<\/p:graphicFrame>/gs) || []).find(f => /compatibility/.test(f)) || null;
  return { label: shape?.missingLabel, frame, off: frame && frame.match(/<a:off x="(\d+)"/)[1],
    vml: pkg.has('ppt/drawings/vmlDrawing1.vml') && pkg.rels('ppt/slides/slide1.xml').some(r => /\/vmlDrawing$/.test(r.type)),
    text: pkg.has('ppt/drawings/legacyDiagramText1.bin'), losses: (pres.losses || []).map(e => ({ id: e.id.split(':')[0], notify: !!e.notify, what: e.what })) };
}
const run = async kind => {
  const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(fs.readFileSync(file).toString('base64'))}, ${JSON.stringify(kind)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

test('an unedited or moved diagram goes back as its frame, with its VML drawing and text', async t => {
  if (!page) return t.skip('Fetch pinned corpus: ' + file);
  const unedited = await run('unedited');
  assert.equal(unedited.label, 'Diagram');
  assert.match(unedited.frame || '', /<com:legacyDrawing[^>]*spid="_x0000_s2056"/);
  assert.ok(unedited.vml && unedited.text, 'the VML drawing and its legacy diagram text are kept');
  assert.deepEqual(unedited.losses, []);
  const moved = await run('moved');
  assert.equal(+moved.off, 5508625 + 20 * 12700, 'the frame moves with the placeholder');
  assert.deepEqual(moved.losses, []);
});

test('an edited diagram says what is saved; a deleted one takes its drawing along', async t => {
  if (!page) return t.skip('Fetch pinned corpus: ' + file);
  const restyled = await run('restyled');
  assert.equal(restyled.frame, null);
  assert.deepEqual(restyled.losses.filter(e => e.id === 'frame'), [{ id: 'frame', notify: true, what: 'This diagram will be saved as an empty rectangle, because it was edited here.' }]);
  const deleted = await run('deleted');
  assert.equal(deleted.frame, null);
  assert.ok(!deleted.vml && !deleted.text);
  assert.ok(!deleted.losses.some(e => e.id === 'unreferenced'));
});
