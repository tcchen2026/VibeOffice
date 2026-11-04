// ActiveX controls are kept while their pictures are unchanged: unedited, moved, resized; deleted ones go quietly;
// an edited picture, a copy or a rotation is saved as the picture, with a notice.
// node --test tests/lectern/activex.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const corpus = path.join(process.env.VO_CORPORA || '/work/corpora', 'powerpoint');
const buttons = path.join(corpus, 'libreoffice__cebccf43c31e__activex_commandbutton.pptx');
const legacy = path.join(corpus, 'poi__fe013ed48fe3__keyframes.pptx');
let page, stop;
before(async () => {
  if (!fs.existsSync(buttons)) return;
  stop = await ensureServer(process.env.BASE);
  page = await openPage('lectern', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

async function save(data, slideIndex, kind, format) {
  const pres = await L.pptx.read(Uint8Array.from(atob(data), c => c.charCodeAt(0))); L.pres = pres;
  const slide = pres.slides[slideIndex], control = slide.shapes.find(s => s.keep?.frame?.label === 'ActiveX control');
  if (kind === 'moved') { control.x += 36; control.y += 18; }
  if (kind === 'resized') control.w += 20;
  if (kind === 'deleted') slide.shapes.splice(slide.shapes.indexOf(control), 1);
  if (kind === 'cropped') control.crop = { l: 0.1, t: 0, r: 0, b: 0 };
  if (kind === 'rotated') control.rot = 15;
  if (kind === 'copied') { const copy = L.model.dup(control); copy.x += 30; slide.shapes.push(copy); }
  const pkg = await L.opc.open(new Uint8Array(await (await L.pptx.write(pres, { format })).arrayBuffer()));
  const part = 'ppt/slides/slide' + (slideIndex + 1) + '.xml', xml = pkg.text(part);
  const vmlPart = pkg.rels(part).find(r => /\/vmlDrawing$/.test(r.type))?.part;
  return {
    controls: (xml.match(/<p:controls>.*<\/p:controls>/s) || [''])[0],
    pictures: (xml.match(/<p:spTree>.*<\/p:spTree>/s)[0].match(/<p:pic>/g) || []).length,
    vml: vmlPart ? pkg.text(vmlPart) : null, activeX: pkg.names.filter(n => /^ppt\/activeX\/activeX\d+\.xml$/.test(n)).length,
    losses: (pres.losses || []).filter(e => !/^(shape-form|content|macros?)\b/.test(e.id)).map(e => ({ id: e.id.split(':')[0], notify: !!e.notify, what: e.what })),
    alarms: (pres.losses || []).filter(e => /^unreferenced:/.test(e.id)).map(e => e.where),
  };
}
const run = async (file, slideIndex, kind, format = 'pptx') => {
  const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(fs.readFileSync(file).toString('base64'))}, ${slideIndex}, ${JSON.stringify(kind)}, ${JSON.stringify(format)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};
const strip = xml => xml.replace(/\s+xmlns:\w+="[^"]*"/g, '');
const names = xml => [...xml.matchAll(/<p:control\b[^>]*\bname="([^"]+)"/g)].map(m => m[1]);

test('unedited, moved and resized controls go back under p:controls with their VML and ActiveX parts', async t => {
  if (!page) return t.skip('Fetch pinned corpus: ' + buttons);
  const unedited = await run(buttons, 0, 'unedited');
  assert.deepEqual(names(unedited.controls), ['CommandButton1', 'CommandButton1', 'CommandButton2', 'CommandButton2', 'CommandButton3', 'CommandButton3']);
  assert.match(unedited.controls, /<mc:Choice Requires="v"><p:control spid="1030"/);
  assert.equal(unedited.pictures, 0, 'no picture beside the kept controls');
  assert.equal(unedited.activeX, 3);
  assert.deepEqual(unedited.losses, []); assert.deepEqual(unedited.alarms, []);
  const moved = await run(buttons, 0, 'moved');
  assert.equal(strip(moved.controls).replace(/<a:off [^>]*>/g, ''), strip(unedited.controls).replace(/<a:off [^>]*>/g, ''), 'only the position changes');
  assert.match(moved.vml, /left:114\.25pt;top:107\.75pt;width:232pt;height:80pt/, 'the VML shape PowerPoint uses moves too; its size stays');
  assert.deepEqual(moved.losses, []);
  const resized = await run(buttons, 0, 'resized');
  assert.equal(names(resized.controls).length, 6);
  assert.match(resized.controls, /<a:ext cx="3209925" cy="1020763"\/>/, 'the picture PowerPoint shows without VML is wider too');
  assert.match(resized.vml, /left:78pt;top:89pt;width:252\.75pt;height:80pt/, 'only the width changes in the VML shape');
  assert.deepEqual(resized.losses, []);
});

test('a deleted control takes its parts; an edited picture, a copy or a rotation is saved as the picture, with a notice', async t => {
  if (!page) return t.skip('Fetch pinned corpus: ' + buttons);
  const deleted = await run(buttons, 0, 'deleted');
  assert.equal(names(deleted.controls).length, 4); assert.equal(deleted.activeX, 2);
  assert.deepEqual(deleted.losses.map(e => [e.id, e.notify]), [['frame', false]]); assert.deepEqual(deleted.alarms, []);
  for (const [kind, what] of [['cropped', 'This ActiveX control will be saved as its picture and stop working, because its picture was edited here.'],
    ['rotated', 'This ActiveX control will be saved as its picture and stop working, because it was rotated or flipped here.'],
    ['copied', 'This copy of an ActiveX control will be saved as its picture and won\'t work.']]) {
    const result = await run(buttons, 0, kind);
    assert.equal(result.pictures, 1, kind);
    assert.equal(names(result.controls).length, kind === 'copied' ? 6 : 4, kind);
    assert.deepEqual(result.losses, [{ id: 'frame', notify: true, what }], kind);
  }
});

test('controls drawn only in VML (a Flash movie, a media player) are read and kept', async t => {
  if (!page || !fs.existsSync(legacy)) return t.skip('Fetch pinned corpus: ' + legacy);
  const result = await run(legacy, 14, 'unedited', 'pptm');
  assert.deepEqual(names(result.controls), ['ShockwaveFlash1']);
  assert.equal(result.pictures, 0);
  assert.deepEqual(result.losses, []);
});
