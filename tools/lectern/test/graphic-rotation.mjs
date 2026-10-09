// Real chart/embedded-object commands and recovery of old rotated model state.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [out] = process.argv.slice(2);
if (!out) throw new Error('Usage: graphic-rotation.mjs OUTPUT');
fs.mkdirSync(out, { recursive: true });
const corpus = path.join(process.env.VO_CORPORA || path.join(os.homedir(), 'corpora'), 'powerpoint');
const stop = await ensureServer(), page = await openPage('lectern', { persistence: false }), reports = [];
try {
  for (const [file, kind] of [['poi__e162cc2c2bf0__bar-chart.pptx', 'chart'], ['libreoffice__12b5c8257d02__ole-emf_min.pptx', 'ole']]) {
    try {
      const result = await page.evaluate(`(${check.toString()})(${JSON.stringify(fs.readFileSync(path.join(corpus, file)).toString('base64'))}, ${JSON.stringify(file)}, ${JSON.stringify(kind)})`);
      for (const [scenario, data] of Object.entries(result.artifacts)) {
        const dir = path.join(out, scenario); fs.mkdirSync(dir, { recursive: true }); fs.writeFileSync(path.join(dir, file), Buffer.from(data, 'base64'));
        reports.push({ file, scenario, status: 'ok', checks: result.checks });
      }
      console.log(kind + ': ' + result.checks + ' checks');
    } catch (error) { reports.push({ file, scenario: kind, status: 'failed', error: error.message }); console.error(error.message); process.exitCode = 1; }
  }
} finally {
  fs.writeFileSync(path.join(out, 'results.jsonl'), reports.map(r => JSON.stringify(r)).join('\n') + '\n');
  fs.writeFileSync(path.join(out, 'browser.json'), JSON.stringify({ errors: page.errors, requests: page.requests }, null, 2));
  if (page.errors.length) process.exitCode = 1;
  await page.close(); stop();
}
async function check(data, file, kind) {
  const { model: M, app: A, ed: E, hist: H, opc: K } = L, artifacts = {}; let checks = 0;
  const assert = (v, why) => { if (!v) throw new Error(why); checks++; };
  VO.opened = () => {};
  H.clear();
  await __corpusHooks.open(new File([Uint8Array.from(atob(data), c => c.charCodeAt(0))], file));
  let slide, shape;
  for (const s of L.pres.slides) M.walk(s.shapes, sh => { if (!shape && (kind === 'chart' ? sh.type === 'chart' : sh.keep?.frame?.label === 'Embedded object')) { shape = sh; slide = s; } return true; });
  assert(shape, 'Missing test object');
  E.goto(L.pres.slides.indexOf(slide), { force: true }); A.view = 'normal'; A.focusArea = 'editor'; E.select([shape.id]); H.clear();
  const initial = JSON.stringify(shape);
  assert(!M.canRotate(shape), 'Rotation is enabled');
  A.rotate(90); A.flip('h');
  assert(JSON.stringify(shape) === initial && !H.undo.length, 'Disabled rotation changed the document');
  async function save(scenario) {
    const blob = await L.pptx.write(L.pres), bytes = new Uint8Array(await blob.arrayBuffer()), pkg = await K.open(bytes);
    for (const part of pkg.names.filter(n => /^ppt\/slides\/[^/]+\.xml$/.test(n))) for (const el of pkg.xml(part).getElementsByTagName('*')) {
      if (el.localName === 'xfrm' && el.namespaceURI === K.NS.p) assert(!['rot', 'flipH', 'flipV'].some(a => el.hasAttribute(a)), 'Graphic frame has illegal orientation');
    }
    let text = ''; for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768));
    artifacts[scenario] = btoa(text);
  }
  if (kind === 'chart') {
    Object.assign(shape, { rot: 45, flipH: true, flipV: true }); await save('old-rotation-state');
    Object.assign(shape, { rot: 0, flipH: false, flipV: false });
  }
  const list = E.find(shape.id).list, group = { id: L.uid('group'), type: 'group', name: 'Rotated group', x: shape.x, y: shape.y, w: shape.w, h: shape.h, rot: 45, kids: [shape] };
  list.splice(list.indexOf(shape), 1, group); E.select([group.id]); H.clear();
  const rotation = shape.rot || 0;
  A.ungroup(); assert((shape.rot || 0) === rotation, 'Ungroup rotated a graphic frame'); await save('ungroup');
  assert(H.doUndo(), 'No ungroup undo'); await save('ungroup-undo');
  assert(H.doRedo(), 'No ungroup redo'); await save('ungroup-redo');
  return { artifacts, checks };
}
