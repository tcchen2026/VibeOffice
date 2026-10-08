// Actual frame edits, detached clipboard transfer and suite draft recovery.
// node tools/lectern/test/frames.mjs OUTPUT_DIR [SCENARIO_REGEX]
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [out, filter = ''] = process.argv.slice(2);
if (!out) throw Error('Usage: frames.mjs OUTPUT_DIR [SCENARIO_REGEX]');
const corpus = path.join(process.env.VO_CORPORA || path.join(os.homedir(), 'corpora'), 'powerpoint');
const samples = [
  ['smartart', 'libreoffice__12c1c1d6e7c5__smartart-linear-rule-vert.pptx', ['geometry', 'copy', 'content', 'ungroup', 'paste', 'draft']],
  ['ole', 'libreoffice__12b5c8257d02__ole-emf_min.pptx', ['geometry', 'copy', 'paste', 'delete']],
  ['ink', 'poi__8436edcd55b8__stress013.pptx', ['geometry', 'copy']],
  ['model3d', 'openxml-sdk__9452348a7f42__3dtestdot.pptx', ['geometry', 'copy']],
  ['chartEx', 'openxml-sdk__254c9ac4ef06__Of16-03.pptx', ['geometry', 'copy']],
];
const cases = samples.flatMap(([kind, file, scenarios]) => scenarios.map(edit => ({ file, kind, edit, scenario: kind + '-' + edit })));
cases.push(...['geometry', 'copy', 'paste-partial', 'delete'].map(edit => ({ file: 'multiple-fallback.pptx', kind: 'chartEx', multi: true, edit, scenario: 'multiple-' + edit })));
const selected = cases.filter(c => new RegExp(filter).test(c.scenario));
fs.mkdirSync(out, { recursive: true });
const results = path.join(out, 'results.jsonl');
const rows = fs.existsSync(results) ? fs.readFileSync(results, 'utf8').trim().split('\n').filter(Boolean).map(JSON.parse).filter(r => !selected.some(c => c.scenario === r.scenario)) : [];
const stop = await ensureServer(), page = await openPage('lectern', { persistence: false });
try {
  for (const c of selected) {
    const row = { ...c, attempted: true }; page.errors.length = 0;
    try {
      const input = path.join(c.multi ? process.env.FRAME_FIXTURES || path.join(path.dirname(out), 'frames-fixtures') : corpus, c.file);
      const result = await page.evaluate(`(${run.toString()})(${JSON.stringify({ ...c, data: fs.readFileSync(input).toString('base64') })})`);
      for (const [state, data] of Object.entries(result.artifacts)) {
        const dir = path.join(out, state === 'saved' ? c.scenario : c.scenario + '-' + state);
        fs.mkdirSync(dir, { recursive: true }); fs.writeFileSync(path.join(dir, c.file), Buffer.from(data, 'base64'));
      }
      delete result.artifacts; Object.assign(row, result);
      if (page.errors.length) throw Error(page.errors.join('\n'));
      if (c.edit === 'draft' || c.edit === 'paste') {
        const shot = await page.send('Page.captureScreenshot', { format: 'png' });
        fs.writeFileSync(path.join(out, c.scenario + '.png'), Buffer.from(shot.data, 'base64'));
      }
    } catch (error) { Object.assign(row, { status: 'failed', error: error.stack }); }
    rows.push(row); fs.writeFileSync(results, rows.map(r => JSON.stringify(r)).join('\n') + '\n');
    console.log(c.scenario + ': ' + row.status + (row.error ? ' ' + row.error : ''));
  }
} finally { await page.close(); stop(); }
if (rows.some(r => r.status !== 'ok')) process.exitCode = 1;

async function run(o) {
  const K = L.opc, M = L.model, A = L.app, H = L.hist, E = L.ed;
  const check = (ok, why) => { if (!ok) throw Error(why); };
  const all = e => e ? [e, ...e.getElementsByTagName('*')] : [];
  const count = root => {
    const result = { smartart: 0, ole: 0, ink: 0, model3d: 0, chartEx: 0 };
    for (const e of all(root)) {
      if (e.localName === 'relIds' && /\/diagram$/.test(e.namespaceURI)) result.smartart++;
      if (e.localName === 'oleObj') result.ole++;
      if (e.localName === 'contentPart') result.ink++;
      if (e.localName === 'model3d') result.model3d++;
      if (e.localName === 'chart' && /\/chartex$/.test(e.namespaceURI)) result.chartEx++;
    }
    return result;
  };
  const stats = pkg => {
    const sum = count(null);
    for (const r of pkg.rels(pkg.main).filter(r => r.type.endsWith('/slide'))) for (const [k, v] of Object.entries(count(pkg.xml(r.part)))) sum[k] += v;
    return sum;
  };
  VO.opened = () => {}; H.clear();
  await __corpusHooks.open(new File([Uint8Array.from(atob(o.data), c => c.charCodeAt(0))], o.file));
  const pres = L.pres, before = stats(pres.pkg), artifacts = {};
  let slide, shape;
  for (const s of pres.slides) {
    M.walk(s.shapes, sh => { if (!shape && sh.keep?.frame && count(K.parse(sh.keep.frame.fragment.xml))[o.kind]) { slide = s; shape = sh; } return true; });
    if (shape) break;
  }
  check(shape, 'No owned ' + o.kind + ' frame');
  const members = o.multi ? slide.shapes.filter(s => s.keep?.frame?.fragment.key === shape.keep.frame.fragment.key) : [shape];
  if (o.multi) check(members.length === 2, 'Expected two fallback members');
  const piece = count(K.parse(shape.keep.frame.fragment.xml)), after = { ...before };
  E.goto(pres.slides.indexOf(slide), { force: true }); A.view = 'normal'; A.focusArea = 'editor'; E.select([shape.id]); H.clear();
  let expected = after;
  async function save(state, blob) {
    blob ||= await L.pptx.write(L.pres);
    const pkg = await K.open(new Uint8Array(await blob.arrayBuffer()));
    const actual = stats(pkg); check(JSON.stringify(actual) === JSON.stringify(expected), 'Frame counts ' + JSON.stringify({ actual, expected, losses: L.pres.losses.filter(x => x.id.startsWith('frame:')) }));
    const data = new Uint8Array(await blob.arrayBuffer()); let text = '';
    for (let i = 0; i < data.length; i += 32768) text += String.fromCharCode(...data.subarray(i, i + 32768));
    artifacts[state] = btoa(text); return pkg;
  }
  if (o.edit.startsWith('paste')) {
    const item = { kind: 'shapes', shapes: [L.clone(shape)], text: '' };
    const packet = JSON.parse(JSON.stringify(L.preserve.clipboard(item)));
    const copied = L.preserve.pasteboard(packet);
    check(copied.shapes[0].keep.source !== pres.pkg.id, 'Source package was not detached');
    const images = []; M.walk(copied.shapes, s => { if (s.media) images.push(s.media); return true; });
    for (const id of images) { check(L.media.has(id), 'Pasted preview missing'); await L.loadImage(L.media.url(id)); }
    A.pasteItem(copied);
  } else if (o.edit === 'ungroup') A.ungroup();
  else {
    H.push('Frame ' + o.edit);
    if (o.edit === 'copy') E.find(shape.id).list.push(...M.dupMany(members));
    else if (o.edit === 'delete') { const list = E.find(shape.id).list; list.splice(list.indexOf(shape), 1); }
    else if (o.edit === 'content') {
      let text; M.walk(shape.kids || [], s => { if (!text && s.tx?.ps[0]?.rs[0]) text = s.tx.ps[0].rs[0]; return true; });
      check(text, 'No editable preview text'); text.t += ' changed';
    } else {
      const bounds = L.unionBounds(members.map(s => ({ x: s.x, y: s.y, w: s.w, h: s.h })));
      for (const shape of members) {
        M.translate(shape, 12 + (shape.x - bounds.x) * .2, 8 + (shape.y - bounds.y) * .1);
        const moved = { x: shape.x, y: shape.y, w: shape.w, h: shape.h };
        shape.w *= 1.2; shape.h *= 1.1;
        if (shape.type === 'group') M.scaleGroup(shape, moved, shape);
      }
    }
  }
  if (['copy', 'paste'].includes(o.edit)) for (const k of Object.keys(after)) after[k] += piece[k];
  if (['content', 'ungroup', 'delete'].includes(o.edit)) for (const k of Object.keys(after)) after[k] -= piece[k];
  await save('saved'); check(H.doUndo(), 'Undo missing'); expected = before; await save('undo');
  check(H.doRedo(), 'Redo missing'); expected = after; await save('redo');
  if (o.edit === 'draft') {
    const info = __corpusHooks.docInfo(L.pres), blob = await __corpusHooks.snapshot(L.pres); await save('draft', blob);
    H.clear(); await __corpusHooks.open(new File([blob], info.draftName || info.name), { draft: true, name: info.name, saved: true, lossState: info.lossState });
    await save('recovered');
  }
  return { status: 'ok', artifacts, before, after, expectedSlides: L.pres.slides.length, losses: L.pres.losses };
}
