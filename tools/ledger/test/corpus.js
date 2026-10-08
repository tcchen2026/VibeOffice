/* Node workbook corpus scenarios, isolated per file so timeouts cannot hide an attempt.
 * node corpus.js DIR RESULTS.jsonl [--save DIR] [--edit | --scenarios save,repeat,...] [--limit N]
 * Browser draft/recovery tests use browser-corpus.mjs; this driver tests the engine. */
const fs = require('node:fs'), path = require('node:path'), { fork } = require('node:child_process');
const { createHash } = require('node:crypto');
const args = process.argv.slice(2);
if (args[0] === '--worker') {
  const L = require('./load.js');
  require('../../../public/ledger/js/layout.js');
  require('../../../public/ledger/js/ops.js');
  require('../../ooxml/scenarios.js');
  (async () => {
    const [file, scenario, save] = args.slice(1), type = path.extname(file).slice(1).toLowerCase();
    const result = { status: 'ok' };
    const write = async (kind, wb) => {
      const bytes = await L.xlsxWrite.bytes(wb, { type: wb.ooxmlFormat || type });
      if (save) { const dir = path.join(save, kind === 'saved' ? scenario : scenario + '-' + kind); fs.mkdirSync(dir, { recursive: true }); fs.writeFileSync(path.join(dir, path.basename(file)), bytes); }
      return bytes;
    };
    try {
      const wb = await L.xlsxRead.read(fs.readFileSync(file));
      const op = corpusScenario('ledger', wb, scenario);
      result.changed = op.changed;
      const bytes = await write('saved', wb);
      result.bytes = bytes.length;
      const reopened = await L.xlsxRead.read(bytes);
      if (scenario === 'repeat') await write('second', reopened);
      if (op.undo) {
        if (!op.undo()) throw new Error('Undo did not restore a transaction');
        await write('undo', wb);
        if (!op.redo()) throw new Error('Redo did not restore a transaction');
        await write('redo', wb);
      }
      if (wb.losses) result.losses = wb.losses;
    } catch (e) { result.status = e.excluded || e.code === 'password' ? 'excluded' : 'failed'; result.error = e.stack || String(e); }
    process.send(result, () => process.disconnect());
  })().catch(e => { process.send({ status: 'failed', error: e.stack }); process.disconnect(); });
} else {
  const opt = (name, fallback) => { const i = args.indexOf(name); return i < 0 ? fallback : args.splice(i, 2)[1]; };
  const flag = name => { const i = args.indexOf(name); return i >= 0 && !!args.splice(i, 1); };
  const save = opt('--save', ''), limit = +opt('--limit', Infinity), timeout = +opt('--timeout', 120000), jobs = +opt('--jobs', 2);
  const edit = flag('--edit'), resume = flag('--resume');
  const scenarios = opt('--scenarios', edit ? 'text,format,insert-row,delete-row,insert-column,delete-column,rename-sheet,geometry,duplicate,delete-object,delete-sheet' : 'save,repeat').split(',');
  if (scenarios.includes('draft')) throw new Error('Use tools/ledger/test/browser-corpus.mjs --scenarios draft for actual snapshot/recovery hooks');
  const [source, output] = args;
  if (!source || !output || args.length !== 2) throw new Error('Usage: corpus.js DIR RESULTS.jsonl [--save DIR] [--edit | --scenarios LIST]');
  const ext = /\.(xlsx|xlsm|xltx|xltm)$/i;
  const files = (fs.statSync(source).isDirectory() ? fs.readdirSync(source).filter(f => ext.test(f)).sort().map(f => path.join(source, f)) : ext.test(source) ? [source] : fs.readFileSync(source, 'utf8').split('\n').filter(Boolean).map(f => path.resolve(path.dirname(source), f))).slice(0, limit);
  fs.mkdirSync(path.dirname(output), { recursive: true });
  const done = new Set();
  if (resume && fs.existsSync(output)) for (const line of fs.readFileSync(output, 'utf8').split('\n')) { try { const r = JSON.parse(line); done.add(r.file + ':' + r.sha256 + ':' + r.scenario); } catch {} }
  if (!resume) fs.writeFileSync(output, '');
  let index = 0, failed = 0, count = 0;
  const tasks = files.flatMap(file => {
    let sha256;
    try { sha256 = createHash('sha256').update(fs.readFileSync(file)).digest('hex'); }
    catch (e) {
      for (const scenario of scenarios) fs.appendFileSync(output, JSON.stringify({ file: path.basename(file), scenario, attempted: true, status: 'failed', error: e.message }) + '\n');
      failed += scenarios.length; count += scenarios.length; return [];
    }
    return scenarios.filter(s => !done.has(path.basename(file) + ':' + sha256 + ':' + s)).map(scenario => ({ file, sha256, scenario }));
  });
  const children = new Set();
  for (const sig of ['SIGINT', 'SIGTERM']) process.once(sig, () => { for (const p of children) p.kill(); process.exit(130); });
  async function worker() {
    while (index < tasks.length) {
      const task = tasks[index++], start = Date.now();
      const result = await new Promise(resolve => {
        const child = fork(__filename, ['--worker', task.file, task.scenario, save], { silent: true });
        children.add(child);
        let message, stderr = '', timedOut = false;
        const timer = setTimeout(() => { timedOut = true; child.kill('SIGKILL'); }, timeout);
        child.on('message', r => { message = r; });
        child.stderr.on('data', d => { stderr = (stderr + d).slice(-4000); });
        child.on('error', e => { message = { status: 'failed', error: e.message }; });
        child.on('close', code => { clearTimeout(timer); children.delete(child); resolve(timedOut ? { status: 'failed', error: 'timeout' } : message || { status: 'failed', error: `exit ${code}: ${stderr}` }); });
      });
      if (result.status === 'failed') failed++;
      fs.appendFileSync(output, JSON.stringify({ ...task, file: path.basename(task.file), attempted: true, ...result, ms: Date.now() - start }) + '\n');
      if (++count % 50 === 0) console.error(`ledger: ${count}/${tasks.length}, ${failed} failures`);
    }
  }
  Promise.all(Array.from({ length: Math.max(1, jobs) }, worker)).then(() => { console.error(`ledger: ${count} attempts, ${failed} failures; ${output}`); if (failed) process.exitCode = 1; });
}
