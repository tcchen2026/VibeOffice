// Real slide operations, retained section/custom-show identities and suite recovery.
// node tools/lectern/test/membership.mjs OUTPUT_DIR (VO_CORPORA overrides ~/corpora)
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [out] = process.argv.slice(2);
if (!out) throw Error('Usage: membership.mjs OUTPUT_DIR');
const corpus = path.join(process.env.VO_CORPORA || path.join(os.homedir(), 'corpora'), 'powerpoint');
const cases = [
  ...['insert', 'copy', 'move', 'delete-section', 'draft'].map(scenario => ({ file: 'libreoffice__a74aecc61390__slide-section-test.pptx', scenario })),
  { file: 'poi__f6381daec5e8__bug65551.pptx', scenario: 'large-empty' },
  { file: 'libreoffice__77bde2711d49__tdf142590.pptx', scenario: 'delete-show' },
  { file: 'libreoffice__bb17471bd09b__tdf131390.pptx', scenario: 'partial-show' },
];
fs.mkdirSync(out, { recursive: true });
const rows = [], stop = await ensureServer(), page = await openPage('lectern', { persistence: false });
try {
  for (const c of cases) {
    const row = { ...c, attempted: true }; page.errors.length = 0;
    try {
      const result = await page.evaluate(`(${run.toString()})(${JSON.stringify({ ...c, data: fs.readFileSync(path.join(corpus, c.file)).toString('base64') })})`);
      for (const [state, data] of Object.entries(result.artifacts)) {
        const folder = path.join(out, state === 'saved' ? c.scenario : c.scenario + '-' + state);
        fs.mkdirSync(folder, { recursive: true }); fs.writeFileSync(path.join(folder, c.file), Buffer.from(data, 'base64'));
      }
      delete result.artifacts;
      Object.assign(row, result);
      if (page.errors.length) throw Error(page.errors.join('\n'));
      if (c.scenario === 'draft') {
        const shot = await page.send('Page.captureScreenshot', { format: 'png' });
        fs.writeFileSync(path.join(out, 'section-draft.png'), Buffer.from(shot.data, 'base64'));
      }
    } catch (error) { Object.assign(row, { status: 'failed', error: error.stack }); }
    rows.push(row); fs.writeFileSync(path.join(out, 'results.jsonl'), rows.map(r => JSON.stringify(r)).join('\n') + '\n');
    console.log(c.scenario + ': ' + row.status + (row.error ? ' ' + row.error : ''));
  }
} finally { await page.close(); stop(); }
if (rows.some(r => r.status !== 'ok')) process.exitCode = 1;

async function run(o) {
  const K = L.opc, M = L.model, A = L.app, H = L.hist, N = K.NS;
  const kids = (e, name) => Array.from(e?.children || []).filter(c => !name || c.localName === name);
  const all = e => e ? [e, ...e.getElementsByTagName('*')] : [];
  const check = (value, message) => { if (!value) throw Error(message); };
  const artifacts = {}, bytes = Uint8Array.from(atob(o.data), c => c.charCodeAt(0));
  VO.opened = () => {};
  H.clear(); // Each case owns its input; do not wait on the previous case's discard dialog.
  await __corpusHooks.open(new File([bytes], o.file));
  const source = L.pres, sectionCount = source.keep.sections.length;
  const maxId = Math.max(255, ...source.slides.map(s => +s.keep.sldId));
  const initial = source.slides.map(s => [s.id, s.section]);
  async function inspect(blob) {
    const pres = L.pres, pkg = await K.open(new Uint8Array(await blob.arrayBuffer()));
    const main = pkg.xml(pkg.main), rels = pkg.rels(pkg.main), ids = kids(kids(main, 'sldIdLst')[0], 'sldId');
    check(ids.length === pres.slides.length, 'Slide count');
    const numbers = ids.map(e => e.getAttribute('id'));
    check(new Set(numbers).size === numbers.length, 'Unique slide IDs');
    const modelIds = new Map(pres.slides.map((s, i) => [s.id, numbers[i]]));
    const partIds = new Map(pres.slides.map((s, i) => [s.id, rels.find(r => r.id === ids[i].getAttributeNS(N.rel, 'id')).part]));
    for (let i = 0; i < pres.slides.length; i++) {
      const s = pres.slides[i];
      if (s.keep?.source === pres.pkg?.id && !s.keep.copy) check(numbers[i] === s.keep.sldId, 'Original ID changed');
      else check(+numbers[i] > maxId, 'New slide ID did not exceed original maximum');
    }
    const sections = all(main).filter(e => e.localName === 'section');
    check(sections.length === sectionCount, 'Empty section definitions were lost');
    for (const section of sections) {
      const actual = kids(kids(section, 'sldIdLst')[0]).map(e => e.getAttribute('id'));
      const expected = pres.slides.filter(s => s.section === section.getAttribute('id')).map(s => modelIds.get(s.id));
      check(JSON.stringify(actual) === JSON.stringify(expected), 'Section membership/order');
      check(actual.every(id => numbers.includes(id)), 'Dangling section slide ID');
    }
    const shows = kids(kids(main, 'custShowLst')[0]);
    for (const record of pres.keep.customShows) {
      const expected = record.slides.filter(id => partIds.has(id)).map(id => partIds.get(id));
      const show = shows.find(e => e.getAttribute('id') === record.id);
      check(!!show === !!expected.length, 'Empty custom show retained or live custom show removed');
      if (show) {
        const actual = kids(kids(show, 'sldLst')[0]).map(e => rels.find(r => r.id === e.getAttributeNS(N.rel, 'id'))?.part);
        check(JSON.stringify(actual) === JSON.stringify(expected), 'Custom-show order/repeated entries');
      }
    }
    const prop = pkg.rels(pkg.main).find(r => r.type.endsWith('/presProps'));
    for (const e of all(prop && pkg.xml(prop.part)).filter(e => e.localName === 'custShow')) check(shows.some(s => s.getAttribute('id') === e.getAttribute('id')), 'Selected custom show is missing');
    return pkg;
  }
  const save = async (state, blob) => {
    blob ||= await L.pptx.write(L.pres); const pkg = await inspect(blob);
    const data = new Uint8Array(await blob.arrayBuffer()); let str = '';
    for (let i = 0; i < data.length; i += 32768) str += String.fromCharCode(...data.subarray(i, i + 32768));
    artifacts[state] = btoa(str); return { blob, pkg };
  };
  A.view = 'normal'; A.focusArea = 'editor'; H.clear();
  if (['insert', 'draft', 'large-empty'].includes(o.scenario)) {
    const before = source.slides[0].section; A.insertSlide('blank', 1); check(source.slides[1].section === before, 'Insert did not join predecessor');
  } else if (o.scenario === 'copy') {
    H.push('Copy across sections'); M.insertSlides(source, 1, [M.dupSlide(source.slides[5])]);
    check(source.slides[1].section === source.slides[0].section, 'Copy retained source section');
  } else if (o.scenario === 'move') {
    const section = source.slides[5].section, moving = source.slides[0]; A.moveSlides([0], 6);
    check(moving.section === section, 'Move did not join destination section');
  } else {
    const selected = o.scenario === 'delete-section' ? [0, 1, 2, 3] : [1];
    A.view = 'sorter'; A.focusArea = 'slides'; A.slideSel = new Set(selected); A.deleteSlides();
  }
  const expectedSlides = L.pres.slides.length;
  await save('saved'); check(H.doUndo(), 'Undo missing');
  check(JSON.stringify(L.pres.slides.map(s => [s.id, s.section])) === JSON.stringify(initial), 'Undo membership');
  await save('undo'); check(H.doRedo(), 'Redo missing'); await save('redo');
  if (o.scenario === 'draft') {
    const info = __corpusHooks.docInfo(L.pres), blob = await __corpusHooks.snapshot(L.pres);
    await save('draft', blob);
    H.clear();
    await __corpusHooks.open(new File([blob], info.draftName || info.name), { draft: true, name: info.name, saved: true, lossState: info.lossState });
    await save('recovered');
  }
  return { status: 'ok', expectedSlides, artifacts, losses: L.pres.losses };
}
