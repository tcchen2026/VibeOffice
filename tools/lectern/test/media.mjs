// Real media objects, actual model transactions, clipboard transport and recovery.
// Usage: node tools/lectern/test/media.mjs FILE_OR_LIST OUTDIR
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [source, out] = process.argv.slice(2);
if (!source || !out) throw new Error('Usage: media.mjs FILE_OR_LIST OUTDIR');
const files = /\.(pptx|pptm|ppsx|ppsm|potx|potm)$/i.test(source) ? [source] : fs.readFileSync(source, 'utf8').trim().split('\n');
fs.mkdirSync(out, { recursive: true });
const results = path.join(out, 'results.jsonl'); fs.writeFileSync(results, '');
const stop = await ensureServer(); let page, destination, failed = 0;
function record(file, scenario, result) {
  for (const [kind, data] of Object.entries(result.artifacts || {})) {
    const directory = path.join(out, kind === 'saved' ? scenario : scenario + '-' + kind);
    fs.mkdirSync(directory, { recursive: true }); fs.writeFileSync(path.join(directory, path.basename(file)), Buffer.from(data, 'base64'));
  }
  delete result.artifacts;
  if (result.status === 'failed') failed++;
  fs.appendFileSync(results, JSON.stringify({ file: path.basename(file), scenario, attempted: true, ...result }) + '\n');
}
try {
  page = await openPage('lectern'); destination = await openPage('lectern');
  for (const file of files) {
    const data = fs.readFileSync(file).toString('base64');
    for (const scenario of ['save', 'geometry', 'shadow', 'crop', 'duplicate', 'delete', 'replace', 'duplicate-slide', 'draft', 'copy']) {
      let result;
      try {
        result = await page.evaluate(`(${exercise.toString()})(${JSON.stringify({ data, name: path.basename(file), scenario })})`);
        if (result.packet) {
          const { sourcePart, shapeId } = result;
          result = { ...await destination.evaluate(`(${paste.toString()})(${JSON.stringify(result.packet)})`), sourcePart, shapeId };
          // Cross-tab transport must not depend on the source page's package/media maps.
          if (destination.errors.length) throw new Error(destination.errors.join('\n'));
        }
        if (page.errors.length) throw new Error(page.errors.join('\n'));
      } catch (error) { result = { status: 'failed', error: error.stack }; }
      record(file, scenario, result);
    }
    console.error(path.basename(file) + ': completed media scenarios');
    await page.close(); await destination.close();
    page = await openPage('lectern'); destination = await openPage('lectern');
  }
} finally { await page?.close(); await destination?.close(); stop(); }
console.error(`${failed} failed attempts; ${results}`); if (failed) process.exitCode = 1;

async function exercise(o) {
  const L = window.L, artifacts = {}, bytes = Uint8Array.from(atob(o.data), c => c.charCodeAt(0));
  const encode = async blob => { const bytes = new Uint8Array(await blob.arrayBuffer()); let s = ''; for (let i = 0; i < bytes.length; i += 32768) s += String.fromCharCode(...bytes.subarray(i, i + 32768)); return btoa(s); };
  const save = async key => { const blob = await L.pptx.write(L.pres); artifacts[key] = await encode(blob); return blob; };
  const pres = await L.pptx.read(bytes); L.app.loadPres(pres, 'Media test', { saved: true }); L.hist.clear();
  let picked;
  pres.slides.forEach((slide, index) => L.model.walk(slide.shapes, shape => { if (!picked && shape.keep?.media) picked = { slide, index, shape }; return true; }));
  if (!picked) return { status: 'excluded', reason: 'No media picture in the input' };
  const { slide, index, shape } = picked; L.ed.goto(index, { force: true }); L.ed.sel = [shape.id]; L.app.focusArea = 'editor';
  const before = await save('before');
  if (o.scenario === 'copy') {
    document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'c', code: 'KeyC', ctrlKey: true, bubbles: true }));
    const event = new ClipboardEvent('copy', { clipboardData: new DataTransfer(), bubbles: true, cancelable: true });
    document.body.dispatchEvent(event);
    const text = event.clipboardData.getData('application/x-lectern');
    if (!text.startsWith('{')) throw new Error('The native copy event did not carry the dependency bundle');
    return { status: 'ok', packet: JSON.parse(text), sourcePart: slide.keep.part, shapeId: shape.keep.identity.ids.find(i => i.kind === 'shape').id };
  }
  if (o.scenario === 'draft') {
    const info = __corpusHooks.docInfo(pres), blob = await __corpusHooks.snapshot(pres);
    artifacts.draft = await encode(blob); VO.opened = () => {}; L.hist.clear();
    await __corpusHooks.open(new File([blob], info.draftName || info.name), { draft: true, name: info.name, saved: true });
    await save('saved');
  } else {
    if (o.scenario !== 'save') L.hist.push('Media ' + o.scenario);
    if (o.scenario === 'geometry') { L.model.translate(shape, 13, 7); shape.w += 9; shape.h += 3; }
    if (o.scenario === 'shadow') shape.shadow = { c: '#345678', a: 0.4, dx: 4, dy: 5, blur: 2 };
    if (o.scenario === 'crop') shape.crop.l = Math.min(0.6, (shape.crop.l || 0) + 0.05);
    if (o.scenario === 'duplicate') L.model.find(slide, shape.id).list.push(L.model.dup(shape));
    if (o.scenario === 'delete') { const list = L.model.find(slide, shape.id).list; list.splice(list.indexOf(shape), 1); slide.anims = slide.anims.filter(a => a.sid !== shape.id); }
    if (o.scenario === 'replace') { shape.type = 'shape'; shape.fill = { t: 'solid', c: '#112233' }; delete shape.media; }
    if (o.scenario === 'duplicate-slide') pres.slides.splice(index + 1, 0, L.model.dupSlide(slide));
    await save('saved');
    if (o.scenario !== 'save') {
      if (!L.hist.doUndo()) throw new Error('Missing undo transaction'); await save('undo');
      if (!L.hist.doRedo()) throw new Error('Missing redo transaction'); await save('redo');
    }
  }
  return { status: 'ok', sourcePart: slide.keep.part, shapeId: shape.keep.identity.ids.find(i => i.kind === 'shape').id, expectedSlides: L.pres.slides.length, expectedBox: [shape.x, shape.y, shape.w, shape.h], artifacts };
}
async function paste(packet) {
  const L = window.L;
  L.app.loadPres(L.model.newPresentation(), 'Pasted media', { saved: true });
  const event = new ClipboardEvent('paste', { clipboardData: new DataTransfer(), bubbles: true, cancelable: true });
  event.clipboardData.setData('application/x-lectern', JSON.stringify(packet)); document.body.dispatchEvent(event);
  let count = 0; for (const slide of L.pres.slides) L.model.walk(slide.shapes, shape => { if (shape.keep?.media) count++; return true; });
  if (count !== 1) throw new Error('Paste did not import the media object');
  const bytes = new Uint8Array(await (await L.pptx.write(L.pres)).arrayBuffer()); let data = '';
  for (let i = 0; i < bytes.length; i += 32768) data += String.fromCharCode(...bytes.subarray(i, i + 32768));
  return { status: 'ok', expectedSlides: L.pres.slides.length, artifacts: { saved: btoa(data) } };
}
