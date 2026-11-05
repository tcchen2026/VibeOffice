// Compare the production CFB reader with saved previous readers, under worker deadlines.
import fs from 'node:fs';
import vm from 'node:vm';
import path from 'node:path';
import { Worker, isMainThread, parentPort, workerData } from 'node:worker_threads';
import { createHash } from 'node:crypto';
if (!isMainThread) {
  const start = performance.now(); let file;
  const bytes = new Uint8Array(fs.readFileSync(workerData.file));
  let source = fs.readFileSync(workerData.reader, 'utf8');
  if (!source.includes('OC.cfb = cfb')) source = source.replace('  /* ================= AES', '  OC.cfb = cfb;\n  /* ================= AES');
  vm.runInThisContext(source);
  const result = { open: 'ok', streams: {}, ms: 0 };
  try { file = globalThis.L.officeCrypto.cfb(bytes); }
  catch (e) { result.open = 'refused'; result.error = { name: e.name, code: e.code, message: e.message }; }
  if (file) for (const name of workerData.streams) {
    try { const data = file.stream(name); result.streams[name] = data ? { status: 'ok', bytes: data.length, sha256: createHash('sha256').update(data).digest('hex') } : { status: 'absent' }; }
    catch (e) { result.streams[name] = { status: 'refused', name: e.name, code: e.code, message: e.message }; }
  }
  result.ms = performance.now() - start; parentPort.postMessage(result);
} else {
  const args = process.argv.slice(2), get = (s, d) => { const i = args.indexOf(s); return i < 0 ? d : args[i + 1]; };
  const root = get('--corpus', '/work/corpora'), out = get('--out', 'tmp/legacy/artifacts/cfb-audit'); fs.mkdirSync(out, { recursive: true });
  const legacy = JSON.parse(fs.readFileSync(path.join(root, 'legacy-manifest.json'))), encrypted = JSON.parse(fs.readFileSync('tmp/legacy/encrypted-ooxml.json'));
  const readers = { before: 'tmp/legacy/crypto-before.js', strict: 'tmp/legacy/crypto-strict.js', current: 'public/common/crypto.js' };
  const inputs = [...legacy.map(x => ({ ...x, group: 'legacy' })), ...encrypted.map(x => ({ ...x, group: 'encrypted' }))];
  const tasks = [], skipped = [];
  for (const item of inputs) {
    const fd = fs.openSync(path.join(root, item.file), 'r'), header = Buffer.alloc(8); fs.readSync(fd, header); fs.closeSync(fd);
    if (header.toString('hex') !== 'd0cf11e0a1b11ae1') { skipped.push(item.file); continue; }
    const streams = item.group === 'encrypted' ? ['EncryptionInfo', 'EncryptedPackage'] : item.format === 'word' ? ['WordDocument', '0Table', '1Table'] : item.format === 'excel' ? ['Workbook', 'Book'] : ['PowerPoint Document', 'Current User'];
    for (const [version, reader] of Object.entries(readers).filter(([name]) => get('--versions', 'before,strict,current').split(',').includes(name))) tasks.push({ item, version, reader, streams: [...streams, '\x05SummaryInformation', '\x05DocumentSummaryInformation'] });
  }
  const output = fs.openSync(path.join(out, 'runs.jsonl'), 'w'), all = []; let next = 0;
  const run = task => new Promise(resolve => {
    const worker = new Worker(new URL(import.meta.url), { workerData: { file: path.join(root, task.item.file), reader: task.reader, streams: task.streams }, resourceLimits: { maxOldGenerationSizeMb: 256 } });
    let done = false;
    const finish = result => { if (done) return; done = true; clearTimeout(timer); worker.terminate(); resolve({ file: task.item.file, sha256: task.item.sha256, group: task.item.group, version: task.version, ...result }); };
    const timer = setTimeout(() => finish({ open: 'timeout' }), 5000);
    worker.on('message', finish); worker.on('error', e => finish({ open: 'crashed', message: e.message }));
    worker.on('exit', code => { if (!done) finish({ open: 'crashed', message: 'worker exit ' + code }); });
  });
  await Promise.all(Array.from({ length: 3 }, async () => { while (next < tasks.length) { const row = await run(tasks[next++]); all.push(row); fs.writeSync(output, JSON.stringify(row) + '\n'); if (all.length % 250 === 0) console.log(`${all.length}/${tasks.length}`); } })); fs.closeSync(output);
  const summary = { documents: inputs.length, nonCFB: skipped.length, versions: {} };
  for (const row of all) { const key = row.group + ':' + row.version, s = summary.versions[key] ||= { total: 0, open: {}, requiredStreamRefusals: 0, optionalStreamRefusals: 0 };
    s.total++; s.open[row.open] = (s.open[row.open] || 0) + 1;
    for (const [name, value] of Object.entries(row.streams || {})) if (value.status === 'refused') s[name.startsWith('\x05') ? 'optionalStreamRefusals' : 'requiredStreamRefusals']++;
  }
  fs.writeFileSync(path.join(out, 'summary.json'), JSON.stringify(summary, null, 2)); console.log(JSON.stringify(summary, null, 2));
}
