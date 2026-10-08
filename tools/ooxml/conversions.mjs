// Real-file conversion notices. Save again after undoing the conversion to
// verify the warning is recomputed, rather than left on the document forever.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { ensureServer, openPage } from './cdp.mjs';
const out = process.argv[2];
if (!out) throw Error('Usage: conversions.mjs OUTPUT_DIR');
const filter = new RegExp(process.argv[3] || '');
const root = process.env.VO_CORPORA || path.join(os.homedir(), 'corpora');
const cases = [
  ['quire', 'word', 'openxml-sdk__8b27091d4e1b__smartTag Nesting.docx', 'wrappers'],
  ['quire', 'word', 'openxml-sdk__31493658a65a__chart.docx', 'chart'],
  ['ledger', 'excel', 'poi__92205f46e8cd__WithChart.xlsx', 'chart'],
  ['lectern', 'powerpoint', 'poi__e162cc2c2bf0__bar-chart.pptx', 'chart'],
  ['lectern', 'powerpoint', 'libreoffice__07eb8b76b2cf__master-slides.pptx', 'notes'],
];
fs.mkdirSync(out, { recursive: true });
const resultPath = path.join(out, 'results.jsonl');
const rows = fs.existsSync(resultPath) ? fs.readFileSync(resultPath, 'utf8').trim().split('\n').filter(Boolean).map(JSON.parse).filter(r => !filter.test(r.scenario)) : [];
const stop = await ensureServer(), pages = new Map();
try {
  for (const [app, folder, file, kind] of cases) {
    if (!filter.test(app + '-' + kind)) continue;
    const row = { app, file, scenario: app + '-' + kind, attempted: true };
    try {
      if (!pages.has(app)) pages.set(app, await openPage(app, { persistence: false }));
      const page = pages.get(app); page.errors.length = 0;
      const result = await page.evaluate(`(${run.toString()})(${JSON.stringify({ app, kind, data: fs.readFileSync(path.join(root, folder, file)).toString('base64') })})`);
      for (const [state, data] of Object.entries(result.artifacts)) {
        const dir = path.join(out, row.scenario + (state === 'saved' ? '' : '-' + state));
        fs.mkdirSync(dir, { recursive: true }); fs.writeFileSync(path.join(dir, file), Buffer.from(data, 'base64'));
      }
      delete result.artifacts; Object.assign(row, result);
      if (page.errors.length) throw Error(page.errors.join('\n'));
    } catch (error) { Object.assign(row, { status: 'failed', error: error.stack }); }
    rows.push(row); console.log(row.scenario + ': ' + row.status + (row.error ? ' ' + row.error : ''));
    fs.writeFileSync(path.join(out, 'results.jsonl'), rows.map(r => JSON.stringify(r)).join('\n') + '\n');
  }
} finally { for (const page of pages.values()) await page.close(); stop(); }
if (rows.some(r => r.status !== 'ok')) process.exitCode = 1;

async function run(o) {
  const bytes = Uint8Array.from(atob(o.data), c => c.charCodeAt(0));
  const doc = o.app === 'quire' ? (await L.docx.read(bytes)).doc : o.app === 'ledger' ? await L.xlsxRead.read(bytes) : await L.pptx.read(bytes);
  const write = () => o.app === 'quire' ? L.docx.write(doc) : o.app === 'ledger' ? L.xlsxWrite.write(doc) : L.pptx.write(doc);
  const check = (v, message) => { if (!v) throw Error(message); }, artifacts = {};
  const save = async state => {
    const blob = await write(), data = new Uint8Array(await blob.arrayBuffer());
    let text = ''; for (let i = 0; i < data.length; i += 32768) text += String.fromCharCode(...data.subarray(i, i + 32768));
    artifacts[state] = btoa(text);
  };
  let restore = () => {}, prefix = o.kind === 'wrappers' ? 'content:' : o.kind + '-edit:';
  if (o.kind === 'wrappers') check(!(doc.losses || []).some(e => e.id.startsWith(prefix)), 'Wrapper warning appeared at open');
  else if (o.kind === 'notes') {
    const slide = doc.slides.find(s => s.keep?.notes); check(slide, 'Missing notes');
    const value = slide.notes; slide.notes += '\nEDITED'; restore = () => { slide.notes = value; };
  } else {
    let target, field;
    if (o.app === 'quire') {
      L.D.walk(doc.main, p => { if (p.t === 'p' && !target) target = p.runs.find(r => r.t === 'chart' && r.src); }); field = 'dirtyChart';
    } else if (o.app === 'ledger') { target = doc.sheets.flatMap(s => s.drawings).find(d => d.kind === 'chart')?.chart; field = 'dirty'; }
    else { for (const slide of doc.slides) L.model.walk(slide.shapes, s => { if (!target && s.chart?.srcId) target = s.chart; return true; }); field = 'edited'; }
    check(target, 'Missing imported chart'); const value = target[field]; target[field] = true;
    restore = () => { if (value === undefined) delete target[field]; else target[field] = value; };
  }
  await save('saved');
  const notices = (doc.losses || []).filter(e => e.id.startsWith(prefix));
  check(notices.length, 'Converted output has no compatibility notice');
  if (o.kind !== 'wrappers') { restore(); await save('undo'); check(!(doc.losses || []).some(e => e.id.startsWith(prefix)), 'Undo left a stale conversion notice'); }
  return { status: 'ok', artifacts, notices };
}
