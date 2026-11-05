// Pinned corpus / hostile input harness. A worker deadline bounds CPU; buffers are capped by the reader.
// node tools/legacy/measure.mjs --out tmp/legacy/artifacts/A [--hostile] [--limit 10]
import fs from 'node:fs';
import path from 'node:path';
import { Worker, isMainThread, parentPort, workerData } from 'node:worker_threads';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { load, root } from './load.mjs';
import { auditShared } from './shared-audit.mjs';
const refusal = e => ['LegacyError', 'CFBError'].includes(e.name);
if (!isMainThread) {
  const B = load(), start = performance.now(); let info, sniff, shared;
  try {
    let bytes = new Uint8Array(fs.readFileSync(workerData.file));
    if (workerData.mutation === 'truncated') bytes = bytes.subarray(0, Math.floor(bytes.length * .57));
    if (workerData.mutation === 'bitflip' && bytes.length) {
      const seed = parseInt(workerData.sha256.slice(0, 8), 16); bytes[seed % bytes.length] ^= 1 << (seed % 8);
    }
    if (workerData.mutation === 'header' && bytes.length > 76) bytes.fill(255, 44, 76);
    sniff = B.sniff(bytes); if (workerData.shared) shared = await auditShared(B, bytes, sniff);
    const converter = B[sniff.format]?.convert;
    // An inspection retains a second complete cell/text tree. A/B references are
    // reusable; converter phases measure the converter itself unless requested.
    if(!converter||workerData.inspect){info=B.inspect(bytes);if(workerData.reference)fs.writeFileSync(workerData.reference,JSON.stringify(info));}
    if (!converter) B.fail('CONVERTER_UNAVAILABLE', 'Conversion for this format has not been implemented');
    const result = await converter(bytes, {extension:path.extname(workerData.file)});
    if (!(result.pkg instanceof Uint8Array) || result.pkg.length > B.limits.bytes || !/^(?:doc|xls|ppt|dot|xlt|pot|pps)[xm]$/.test(result.type)) throw new Error('Invalid converter result');
    parentPort.postMessage({ status: 'converted', sniff, shared, features: info?.features, stats:result.stats, ms: performance.now() - start, losses: result.losses, type: result.type, pkg: result.pkg }, [result.pkg.buffer]);
  } catch (e) { parentPort.postMessage({ status: refusal(e) ? 'refused' : 'crashed', sniff, shared, features: info?.features, code: e.code, reason: e.message, where: e.where, stack: refusal(e) ? undefined : e.stack, ms: performance.now() - start }); }
} else {
  const args = process.argv.slice(2), opt = (name, fallback) => { const i = args.indexOf(name); return i < 0 ? fallback : args[i + 1]; };
  const corpus = path.resolve(opt('--corpus', '/work/corpora')), out = path.resolve(opt('--out', 'tmp/legacy/artifacts/measure'));
  const manifest = JSON.parse(fs.readFileSync(path.join(corpus, 'legacy-manifest.json'), 'utf8'));
  const files=opt('--files','')?new Set(JSON.parse(fs.readFileSync(opt('--files',''),'utf8'))):null;
  const selected = manifest.filter(e => (!files||files.has(e.file))&&(!opt('--format', '') || e.format === opt('--format', ''))&&(!opt('--match','')||e.file.includes(opt('--match','')))).slice(0, +opt('--limit', Infinity));
  fs.mkdirSync(path.join(out, 'references'), { recursive: true }); fs.mkdirSync(path.join(out, 'packages'), { recursive: true });
  const run = (item, mutation) => new Promise(resolve => {
    const worker = new Worker(new URL(import.meta.url), { workerData: { file: path.join(corpus, item.file), sha256: item.sha256, mutation, shared: args.includes('--shared'),inspect:args.includes('--inspect'),
      reference: mutation ? null : path.join(out, 'references', item.sha256 + '.json') }, resourceLimits: { maxOldGenerationSizeMb: +opt('--memory',512), stackSizeMb: 8 } });
    let done = false;
    const finish = result => { if (done) return; done = true; clearTimeout(timer); worker.terminate(); resolve({ file: item.file, sha256: item.sha256, size: item.size, mutation, ...result }); };
    const timer = setTimeout(() => finish({ status: 'crashed', code: 'TIMEOUT', reason: 'Worker deadline exceeded' }), +opt('--timeout', 30000));
    worker.on('message', result => {
      if (result.pkg) {
        const file = path.join(out, 'packages', item.sha256 + (mutation ? '-' + mutation : '') + '.' + result.type);
        fs.writeFileSync(file, result.pkg); result.outputHash = createHash('sha256').update(result.pkg).digest('hex'); delete result.pkg;
        const check = spawnSync('python3', ['-I', path.join(root, 'tools/ooxml/package.py'), 'check', file], { encoding: 'utf8', timeout: 60000, maxBuffer: 8 * 1024 * 1024 });
        result.packageCheck = { status: check.status, report: check.stdout, error: check.stderr };
        if (check.status !== 0) result.status = 'invalid-package';
      }
      finish(result);
    });
    worker.on('error', e => finish({ status: 'crashed', code: 'WORKER_ERROR', reason: e.message }));
    worker.on('exit', code => { if (!done) finish({ status: 'crashed', code: 'WORKER_EXIT', reason: String(code) }); });
  });
  const all = [], output = fs.openSync(path.join(out, 'runs.jsonl'), 'w'); let next = 0;
  const tasks = selected.flatMap(item => (args.includes('--hostile') ? [null, 'truncated', 'bitflip', 'header'] : [null]).map(mutation => ({ item, mutation })));
  await Promise.all(Array.from({ length: +opt('--workers', 2) }, async () => {
    while (next < tasks.length) {
      const { item, mutation } = tasks[next++], result = await run(item, mutation); all.push(result); fs.writeSync(output, JSON.stringify(result) + '\n');
      if (all.length % 100 === 0) console.log(`${all.length}/${tasks.length}`);
    }
  })); fs.closeSync(output);
  const summary = { total: all.length, originals: all.filter(r => !r.mutation).length, status: {}, versions: {}, reasons: {}, largest: {} };
  for (const r of all) { summary.status[r.status] = (summary.status[r.status] || 0) + 1; if (r.mutation) continue;
    const version = (r.sniff?.format || 'unidentified') + ':' + (r.sniff?.version || r.code); summary.versions[version] = (summary.versions[version] || 0) + 1;
    if (r.code) summary.reasons[r.code] = (summary.reasons[r.code] || 0) + 1;
    const format = r.sniff?.format || 'unknown'; if (!summary.largest[format] || summary.largest[format].size < r.size) summary.largest[format] = r;
  }
  fs.writeFileSync(path.join(out, 'summary.json'), JSON.stringify(summary, null, 2) + '\n'); console.log(JSON.stringify(summary, null, 2));
  if (all.some(r => ['crashed', 'invalid-package'].includes(r.status))) process.exitCode = 1;
}
