// Recover misplaced page/line breaks without keeping invalid Word structure.
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../ooxml/cdp.mjs';
const [corpus, output] = process.argv.slice(2);
if (!corpus || !output) throw new Error('Usage: recovered-breaks.mjs CORPUS OUTPUT');
const file = 'libreoffice__145e32d57e82__tdf108714.docx';
fs.mkdirSync(output, { recursive: true });
const stop = await ensureServer(), page = await openPage('quire', { persistence: false }), reports = [];
try {
  const result = await page.evaluate(`(${check.toString()})(${JSON.stringify(fs.readFileSync(path.join(corpus, file)).toString('base64'))})`);
  for (const [scenario, value] of Object.entries(result.artifacts)) {
    const dir = path.join(output, scenario); fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, file), Buffer.from(value, 'base64'));
    reports.push({ file, scenario, status: 'ok', checks: result.checks });
  }
  if (page.errors.length) throw new Error(page.errors.join('\n'));
  console.log(result.checks + ' recovered-break checks');
} catch (error) {
  reports.push({ file, scenario: 'breaks', status: 'failed', error: error.message }); process.exitCode = 1;
  console.error(error.message);
} finally {
  fs.writeFileSync(path.join(output, 'results.jsonl'), reports.map(r => JSON.stringify(r)).join('\n') + '\n');
  fs.writeFileSync(path.join(output, 'browser.json'), JSON.stringify({ errors: page.errors, requests: page.requests }, null, 2));
  await page.close(); stop();
}
async function check(data) {
  const { D, O, opc: K } = L, artifacts = {}; let checks = 0;
  const assert = (ok, message) => { if (!ok) throw new Error(message); checks++; };
  const bytes = Uint8Array.from(atob(data), c => c.charCodeAt(0));
  let doc = (await L.docx.read(bytes)).doc; L.app.loadDoc(doc, 'Recovered breaks', { saved: true });
  const paragraphs = () => { const out = []; D.walk(doc.main, p => { if (p.t === 'p') out.push(p); }); return out; };
  const breaks = () => paragraphs().flatMap(p => p.runs.filter(r => r.t === 'br'));
  const notice = () => doc.losses?.some(e => e.id.startsWith('break-placement:'));
  assert(breaks().length === 7, 'Not all misplaced breaks were recovered');
  assert(breaks().filter(r => r.type === 'page').length === 3, 'The page breaks changed type');
  assert(!notice(), 'A recovery notice was added before preparing a save');
  const save = async (scenario, blob) => {
    blob ||= await L.docx.write(doc); const a = new Uint8Array(await blob.arrayBuffer()); let s = '';
    for (let i = 0; i < a.length; i += 32768) s += String.fromCharCode(...a.subarray(i, i + 32768));
    artifacts[scenario] = btoa(s);
    const pkg = await K.open(a), tree = pkg.xml(pkg.main);
    const nodes = Array.from(tree.getElementsByTagName('*')).filter(e => e.localName === 'br');
    assert(nodes.length === breaks().length, 'Saved break count differs from the model');
    assert(nodes.every(e => e.parentNode.localName === 'r'), 'Saved break is still outside a text run');
  };
  await save('save'); assert(notice(), 'Missing recovery notice');
  const p = paragraphs().find(p => p.runs.some(r => r.t === 'text'));
  D.tx('Edit text', () => O.insertText(D.pos(p, 0), 'Edited ')); await save('text');
  D.tx('Remove recovered breaks', () => { for (const p of paragraphs()) { D.touch(p); p.runs = p.runs.filter(r => !r.recoveredPlacement); } });
  await save('remove'); assert(!notice(), 'Deleted breaks left a stale recovery notice');
  assert(D.undo(), 'Missing undo'); await save('undo'); assert(notice(), 'Undo lost the recovery notice');
  assert(D.redo(), 'Missing redo'); await save('redo'); assert(!notice(), 'Redo left a stale notice');
  assert(D.undo(), 'Missing restore');
  const h = __corpusHooks, draft = await h.snapshot(doc), info = h.docInfo(doc); await save('draft', draft);
  VO.opened = () => {};
  await h.open(new File([draft], 'breaks.docx'), { draft: true, name: info.name, saved: true, lossState: info.lossState });
  doc = h.current(); await save('recovered');
  assert(breaks().length === 7 && notice(), 'Draft recovery lost breaks or the pending conversion notice');
  return { artifacts, checks };
}
