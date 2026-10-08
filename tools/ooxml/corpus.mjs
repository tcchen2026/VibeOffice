// CDP corpus driver. Keeps user tabs open and records every attempt.
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { ensureServer, openPage } from './cdp.mjs';

export async function corpus(app, argv = process.argv.slice(2)) {
  const args = argv.slice();
  const opt = (name, fallback) => { const i = args.indexOf(name); return i < 0 ? fallback : args.splice(i, 2)[1]; };
  const flag = name => { const i = args.indexOf(name); return i >= 0 && !!args.splice(i, 1); };
  const save = opt('--save', null), limit = +opt('--limit', Infinity), timeout = +opt('--timeout', 120000);
  const edits = flag('--edit'), resume = flag('--resume');
  const editScenarios = { quire: 'text,format,geometry,duplicate,delete-object', ledger: 'text,format,geometry,duplicate,delete-object,insert-row,delete-row,insert-column,delete-column,rename-sheet,delete-sheet', lectern: 'text,format,geometry,duplicate,delete-object,insert-slide,delete-slide' };
  const scenarios = opt('--scenarios', edits ? editScenarios[app] : 'save,repeat').split(',');
  const [source, output] = args;
  if (!source || !output || args.length !== 2) throw new Error('Usage: corpus SOURCE RESULTS.jsonl [--save DIR] [--edit | --scenarios save,repeat,text,...] [--limit N] [--resume]');
  const ext = app === 'quire' ? /\.(docx|docm|dotx|dotm)$/i : app === 'ledger' ? /\.(xlsx|xlsm|xltx|xltm)$/i : /\.(pptx|pptm|ppsx|ppsm|potx|potm)$/i;
  const files = (fs.statSync(source).isDirectory() ? fs.readdirSync(source).filter(f => ext.test(f)).sort().map(f => path.join(source, f)) : ext.test(source) ? [source] : fs.readFileSync(source, 'utf8').split('\n').filter(Boolean).map(f => path.resolve(path.dirname(source), f))).slice(0, limit);
  const done = new Set();
  if (resume && fs.existsSync(output)) for (const line of fs.readFileSync(output, 'utf8').split('\n')) { try { const r = JSON.parse(line); done.add(r.file + ':' + r.sha256 + ':' + r.scenario); } catch { /* incomplete line is retried */ } }
  fs.mkdirSync(path.dirname(output), { recursive: true });
  if (!resume) fs.writeFileSync(output, '');
  const stopServer = await ensureServer();
  let page, n = 0, failed = 0;
  const close = async () => { await page?.close(); stopServer(); };
  const interrupt = async () => { await close(); process.exit(130); };
  process.once('SIGINT', interrupt); process.once('SIGTERM', interrupt);
  const helper = fs.readFileSync(fileURLToPath(new URL('./scenarios.js', import.meta.url)), 'utf8');
  const fresh = async () => { await page?.close(); page = await openPage(app, { timeout, persistence: false }); await page.evaluate(helper); };
  try {
    await fresh();
    for (const file of files) {
      let bytes, sha256;
      try { bytes = fs.readFileSync(file); sha256 = createHash('sha256').update(bytes).digest('hex'); }
      catch (e) {
        for (const scenario of scenarios) fs.appendFileSync(output, JSON.stringify({ file: path.basename(file), scenario, attempted: true, status: 'failed', error: e.message }) + '\n');
        n += scenarios.length; failed += scenarios.length; continue;
      }
      for (const scenario of scenarios) {
        const base = { file: path.basename(file), sha256, scenario, attempted: true };
        if (done.has(base.file + ':' + sha256 + ':' + scenario)) continue;
        const t = Date.now();
        let result;
        page.errors.length = 0;
        try {
          result = await page.evaluate(`(${run.toString()})(${JSON.stringify({ app, scenario, data: bytes.toString('base64'), name: path.basename(file), save: !!save })})`);
          if (result.artifacts) {
            for (const [kind, content] of Object.entries(result.artifacts)) {
              const dir = path.join(save, kind === 'saved' ? scenario : scenario + '-' + kind);
              fs.mkdirSync(dir, { recursive: true });
              fs.writeFileSync(path.join(dir, path.basename(file)), Buffer.from(content, 'base64'));
            }
            delete result.artifacts;
          }
          if (page.errors.length) { result.consoleErrors = page.errors.slice(); result.status = 'failed'; }
        } catch (e) { result = { status: 'failed', error: e.message }; await fresh(); }
        if (result.status === 'failed') failed++;
        fs.appendFileSync(output, JSON.stringify({ ...base, ...result, ms: Date.now() - t }) + '\n');
        if (++n % 25 === 0) { console.error(`${app}: ${n} attempts, ${failed} failures`); await fresh(); }
      }
    }
  } finally { await close(); process.removeListener('SIGINT', interrupt); process.removeListener('SIGTERM', interrupt); }
  console.error(`${app}: ${n} attempts, ${failed} failures; ${output}`);
  if (failed) process.exitCode = 1;
}

async function run(o) {
  const L = window.L, result = { status: 'ok' }, artifacts = {};
  const bytes = Uint8Array.from(atob(o.data), c => c.charCodeAt(0));
  const format = o.name.split('.').pop().toLowerCase();
  const read = async data => o.app === 'quire' ? (await L.docx.read(data)).doc : o.app === 'ledger' ? L.xlsxRead.read(data) : L.pptx.read(data);
  const write = doc => {
    const type = doc.ooxmlFormat || format;
    return o.app === 'quire' ? L.docx.write(doc, { format: type }) : o.app === 'ledger' ? L.xlsxWrite.write(doc, { type }) : L.pptx.write(doc, { format: type });
  };
  const encode = async blob => { const b = new Uint8Array(await blob.arrayBuffer()); let s = ''; for (let i = 0; i < b.length; i += 32768) s += String.fromCharCode(...b.subarray(i, i + 32768)); return btoa(s); };
  const save = async (kind, doc) => { const blob = await write(doc); if (o.save) artifacts[kind] = await encode(blob); return blob; };
  try {
    let doc = await read(bytes);
    if (o.app === 'lectern' && doc.pkg) {
      const list = Array.from(doc.pkg.xml(doc.pkg.main).children).find(e => e.localName === 'sldIdLst');
      const original = list ? Array.from(list.children).filter(e => e.localName === 'sldId').length : 0;
      result.slides = { original, opened: doc.slides.length };
      if (original !== doc.slides.length) throw new Error(`Opened ${doc.slides.length} of ${original} source slides`);
    }
    if (o.app === 'quire') L.D.doc = doc; else if (o.app === 'lectern') { L.pres = doc; L.hist.clear(); }
    const operation = corpusScenario(o.app, doc, o.scenario);
    result.changed = operation.changed;
    if (operation.expectedSlides !== undefined) result.expectedSlides = operation.expectedSlides;
    const first = await save('saved', doc);
    result.bytes = first.size;
    const reread = await read(new Uint8Array(await first.arrayBuffer()));
    if (result.slides) {
      result.slides.saved = reread.slides.length;
      if (doc.slides.length !== reread.slides.length) throw new Error(`Saved ${reread.slides.length} of ${doc.slides.length} model slides`);
    }
    if (o.scenario === 'repeat') await save('second', reread);
    if (operation.undo) {
      if (!operation.undo()) throw new Error('Undo did not restore a transaction');
      await save('undo', doc);
      if (!operation.redo()) throw new Error('Redo did not restore a transaction');
      await save('redo', doc);
    }
    if (o.scenario === 'draft') {
      // Actual hooks, with recent-file bookkeeping disabled only in this test tab.
      VO.opened = () => {};
      await __corpusHooks.open(new File([bytes], o.name));
      doc = __corpusHooks.current();
      const draft = await __corpusHooks.snapshot(doc), info = __corpusHooks.docInfo(doc);
      result.draftName = info.draftName || info.name;
      result.draftType = draft.type;
      if (o.save) artifacts.draft = await encode(draft);
      await __corpusHooks.open(new File([draft], result.draftName), { draft: true, name: info.name, saved: true, lossState: info.lossState });
      const recovered = __corpusHooks.current();
      doc = o.app === 'ledger' ? recovered.wb : recovered;
      await save('recovered', doc);
    }
    if (doc.losses) result.losses = doc.losses;
    if (o.save) result.artifacts = artifacts;
  } catch (e) {
    result.status = e.excluded || e.code === 'password' ? 'excluded' : 'failed';
    result.error = e.stack || String(e);
  }
  return result;
}
