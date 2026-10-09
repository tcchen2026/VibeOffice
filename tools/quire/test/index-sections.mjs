// Insert a multi-column index in an existing section, then exercise history and drafts.
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [output] = process.argv.slice(2);
if (!output) throw new Error('Usage: index-sections.mjs OUTPUT');
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), page = await openPage('quire', { persistence: false }), results = [];
try {
  for (const following of [false, true]) {
    const file = following ? 'index-middle-section.docx' : 'index-final-section.docx';
    try {
      const result = await page.evaluate(`(${check.toString()})(${following})`);
      for (const [scenario, data] of Object.entries(result.artifacts)) {
        const dir = path.join(output, scenario); fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(path.join(dir, file), Buffer.from(data, 'base64'));
        if (scenario !== 'original') results.push({ file, scenario, status: 'ok', checks: result.checks });
      }
      console.log(file + ': ' + result.checks + ' checks');
    } catch (error) { results.push({ file, scenario: 'index', status: 'failed', error: error.message }); process.exitCode = 1; console.error(error.message); }
  }
} finally {
  fs.writeFileSync(path.join(output, 'results.jsonl'), results.map(r => JSON.stringify(r)).join('\n') + '\n');
  fs.writeFileSync(path.join(output, 'browser.json'), JSON.stringify({ errors: page.errors, requests: page.requests }, null, 2));
  if (page.errors.length) process.exitCode = 1;
  await page.close(); stop();
}
async function check(following) {
  const { D, fields: F, opc: K } = L, artifacts = {}; let checks = 0;
  const assert = (condition, message) => { if (!condition) throw new Error(message); checks++; };
  let doc = D.newDoc();
  doc.main.blocks = ['Previous section', 'Index position', 'Section remainder', 'Following section'].map(s => D.para([D.text(s)]));
  const section = type => Object.assign(D.defaultSect(), { type });
  doc.main.blocks[0].sect = section('nextPage');
  if (following) doc.main.blocks[2].sect = section('evenPage');
  doc.sect = section(following ? 'oddPage' : 'evenPage');
  L.app.loadDoc(doc, 'Index sections', { saved: true });
  const encode = async blob => {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let text = '';
    for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768));
    return btoa(text);
  };
  const save = async (scenario, blob) => {
    blob ||= await L.docx.write(doc); artifacts[scenario] = await encode(blob);
    const pkg = await K.open(new Uint8Array(await blob.arrayBuffer()));
    return Array.from(pkg.xml(pkg.main).getElementsByTagName('*')).filter(e => e.localName === 'sectPr').map(e => {
      const type = Array.from(e.children).find(c => c.localName === 'type');
      return type?.getAttribute('w:val') || 'nextPage';
    });
  };
  const original = await save('original');
  const pos = D.pos(doc.main.blocks[1], 5); L.ed.sel = { a: pos, f: pos };
  F.insertGenerated('INDEX \\c "2"', { columns: 2 });
  const wanted = ['nextPage', 'evenPage', 'continuous', 'continuous', ...(following ? ['oddPage'] : [])];
  const inserted = await save('index');
  assert(JSON.stringify(inserted) === JSON.stringify(wanted), 'Index moved the section start after its columns: ' + JSON.stringify(inserted));
  const history = D.hist.undo.length;
  assert(history > 0, 'No index undo');
  for (let i = 0; i < history; i++) assert(D.undo(), 'Missing index/page-number undo');
  const undone = await save('index-undo');
  assert(JSON.stringify(undone) === JSON.stringify(original), 'Undo changed the original section starts: ' + JSON.stringify(undone) + ' expected ' + JSON.stringify(original));
  for (let i = 0; i < history; i++) assert(D.redo(), 'Missing index/page-number redo');
  assert(JSON.stringify(await save('index-redo')) === JSON.stringify(wanted), 'Redo moved a section start');
  const snapshot = await window.__corpusHooks.snapshot(doc), info = window.__corpusHooks.docInfo(doc);
  await save('draft', snapshot); VO.opened = () => {};
  await window.__corpusHooks.open(new File([snapshot], 'index.docx'), { draft: true, name: info.name, saved: true, lossState: info.lossState });
  doc = window.__corpusHooks.current();
  assert(JSON.stringify(await save('recovered')) === JSON.stringify(wanted), 'Recovery changed section starts');
  return { artifacts, checks };
}
