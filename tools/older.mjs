// Run today's tests against older code: public/ as it was at a commit, served on a port of its own (BASE). A fix's
// regression test must fail there (todos/UI_MOUSE_LOGIC_IMPROVE.md, Working rules; docs/testing.md).
//
//   node tools/older.mjs [REV] [-- <node --test arguments>]
//   node tools/older.mjs 1035247~1 -- tests/ui/painter.test.mjs
//   node tools/older.mjs -- --test-name-pattern '^Q1 ' tests/ui/quire.test.mjs
//
// REV defaults to HEAD while public/ has uncommitted changes (a fix not committed yet), else to HEAD~1 (the fix is
// the last commit). The copy goes to tmp/older/<commit>/ (kept for the next run) and is served by tools/serve.py on
// the first free port from 8761. Without test arguments it runs tests/ui/*.test.mjs, one file at a time.
import fs from 'node:fs';
import net from 'node:net';
import path from 'node:path';
import { spawn, execFileSync } from 'node:child_process';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const git = (...a) => execFileSync('git', a, { cwd: root, encoding: 'utf8' }).trim();
const argv = process.argv.slice(2);
const cut = argv.indexOf('--');
const [revArg] = cut < 0 ? argv : argv.slice(0, cut);
let testArgs = cut < 0 ? [] : argv.slice(cut + 1);
if (!testArgs.some((a) => !a.startsWith('-'))) {
  const dir = path.join(root, 'tests', 'ui');
  testArgs = [...testArgs, ...fs.readdirSync(dir).filter((f) => f.endsWith('.test.mjs')).sort().map((f) => path.join('tests', 'ui', f))];
}

const rev = revArg || (git('status', '--porcelain', '--', 'public') ? 'HEAD' : 'HEAD~1');
const sha = git('rev-parse', '--short', rev + '^{commit}');
const dir = path.join(root, 'tmp', 'older', sha);
if (!fs.existsSync(path.join(dir, 'public'))) {
  fs.mkdirSync(dir, { recursive: true });
  execFileSync('sh', ['-c', `git archive ${sha} public | tar -x -C '${dir}'`], { cwd: root, stdio: 'inherit' });
}

const listening = (port) => new Promise((r) => { const s = net.connect(port, '127.0.0.1', () => { s.end(); r(true); }).on('error', () => r(false)); });
let port = 8761;
while (await listening(port)) port++;
const server = spawn('python3', [path.join(root, 'tools', 'serve.py'), String(port), path.join(dir, 'public')], { stdio: 'ignore' });
for (let i = 0; i < 50 && !(await listening(port)); i++) await new Promise((r) => setTimeout(r, 100));
const BASE = `http://127.0.0.1:${port}`;
console.log(`older code: public/ at ${git('log', '-1', '--format=%h %s', sha)}\n  served on ${BASE}; today's tests: ${testArgs.join(' ')}\n`);

const run = spawn(process.execPath, ['--test', '--test-concurrency=1', ...testArgs], { cwd: root, stdio: 'inherit', env: { ...process.env, BASE, VO_TEST_PUBLIC: path.join(dir, 'public') } });
const stop = () => server.kill();
for (const sig of ['SIGINT', 'SIGTERM']) process.on(sig, () => { run.kill(sig); stop(); });
run.on('exit', (code) => { stop(); process.exit(code ?? 1); });
