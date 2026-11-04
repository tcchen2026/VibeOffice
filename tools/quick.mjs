// The quick check after a change (docs/testing.md): the tests that cover the files changed, chosen by the rules below,
// in a few minutes; the whole of tests/ui/ waits for a release (tools/ui-full.mjs). Always the smoke test when the
// apps changed (tests/ui/smoke.test.mjs); the Node tests of each part touched; a touched app's own UI tests; focus and
// drag when the shared event, menu and dialog code changed; and every test file changed itself. Fast tests first.
//
//   node tools/quick.mjs              the uncommitted changes, or the last commit when there are none
//   node tools/quick.mjs <rev>        everything changed since <rev> (e.g. origin/main), committed or not
//   node tools/quick.mjs --dry        list what would run, and the time it takes
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync, execFileSync } from 'node:child_process';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const git = (...a) => execFileSync('git', a, { cwd: root, encoding: 'utf8' }).trim();
const args = process.argv.slice(2), dry = args.includes('--dry'), rev = args.find((a) => !a.startsWith('--'));

/* the files changed */
const lines = (s) => s.split('\n').filter(Boolean);
let changed, what;
if (rev) { changed = [...lines(git('diff', '--name-only', rev)), ...lines(git('ls-files', '--others', '--exclude-standard'))]; what = `since ${rev}`; }
else {
  changed = [...lines(git('diff', '--name-only', 'HEAD')), ...lines(git('ls-files', '--others', '--exclude-standard'))];
  what = 'uncommitted';
  if (!changed.length) { changed = lines(git('diff', '--name-only', 'HEAD~1', 'HEAD')); what = `the last commit, ${git('log', '-1', '--format=%h %s')}`; }
}
changed = [...new Set(changed)];

/* rules: changed path → test files. Seconds are measured run times (2026-10-14), for the estimate. */
const SECONDS = { 'tests/ui/smoke.test.mjs': 20, 'tests/ui/quire.test.mjs': 175, 'tests/ui/ledger.test.mjs': 130, 'tests/ui/lectern.test.mjs': 130,
  'tests/ui/focus.test.mjs': 265, 'tests/ui/focus.test.mjs :: classic': 90, 'tests/ui/drag.test.mjs': 45, 'tests/ui/gestures.test.mjs': 480, 'tests/ui/painter.test.mjs': 40, 'tests/ui/textmap.test.mjs': 165,
  'tests/ledger/*.test.js': 5, 'tests/quire/*.test.js': 1, 'tests/quire/lines.test.mjs': 5, 'tests/from-issues/*/*.test.mjs': 1, 'tests/scout.test.mjs': 190, 'tests/desktop/host.test.mjs': 32, 'tests/sw.test.mjs': 10,
  'tests/ribbon.test.mjs': 5, 'tests/search.test.mjs': 4, 'tests/ooxml/check.test.mjs': 3, 'tests/i18n.test.mjs': 1, 'tests/fonts.test.mjs': 1, 'tests/locales.test.mjs': 15,
  'tests/hostile/export-html.test.mjs': 6, 'tests/hostile/helpers.test.mjs': 7, 'tests/csp.test.mjs': 1 };
const RULES = [
  [/^tools\/win_test_server\/|^tests\/win_test_server\//, ['tests/win-test-server.test.mjs']],
  [/^public\/(quire|ledger|lectern|common)\/|^public\/[^/]+\.(html|js|css)$/, ['tests/ui/smoke.test.mjs', 'tests/licenses.test.mjs', 'tests/from-issues/*/*.test.mjs', 'tests/i18n.test.mjs']],   // and the reported issues' tests, the English catalog
  [/^public\/ledger\/|^tools\/ledger\//, ['tests/ledger/*.test.js', 'tests/ui/ledger.test.mjs']],
  [/^public\/quire\//, ['tests/quire/*.test.js', 'tests/ui/quire.test.mjs']],
  [/^public\/quire\/js\/(render|portions|layout|drawing)\.js$|^tests\/lib\/lines\.mjs$/, ['tests/quire/lines.test.mjs']],   // line heights against Word and LibreOffice
  [/^public\/lectern\//, ['tests/ooxml/lectern-review.test.mjs', 'tests/ui/lectern.test.mjs']],
  [/^public\/common\/(core|dml|numfmt|xml|opc|opc-order|zip|sha|crypto|metafile|geometry|charts)\.js$/, ['tests/ledger/*.test.js', 'tests/quire/*.test.js']],
  [/^public\/common\/(ui|ribbon|search)\.js$/, ['tests/ui/focus.test.mjs :: classic', 'tests/ui/drag.test.mjs']],   // focus in the Classic look
  [/^public\/common\/(ui\.css|looks\.css)$|^public\/[^/]+\/css\//, ['tests/looks.test.mjs']],
  [/^public\/common\/ribbon\.js$|^public\/[^/]+\/js\/ribbon\.js$/, ['tests/ribbon.test.mjs']],
  [/^public\/common\/search\.js$/, ['tests/search.test.mjs']],
  [/^public\/common\/smartart/, ['tests/smartart.test.mjs']],
  [/^public\/common\/zip\.js$|^public\/sw\.js$/, ['tests/sw.test.mjs']],
  [/^public\/common\/crypto\.js$/, ['tests/crypto.test.mjs']],
  [/^public\/.*\.js$/, ['tests/insecure.test.mjs']],   // a page served over plain http: no randomUUID or crypto.subtle
  [/^public\/common\/spell\.js$/, ['tests/autocorrect.test.mjs']],
  [/^public\/common\/geometry\.js$|^public\/ledger\/js\/render\.js$|^tools\/patterns\//, ['tests/patterns.test.mjs']],
  [/^public\/common\/(xml|opc|opc-order)\.js$/, ['tests/ooxml/opc.test.mjs', 'tests/ooxml/check.test.mjs']],
  [/^tools\/ooxml\/.*\.py$/, ['tests/ooxml/check.test.mjs', 'tests/ooxml/audit.test.mjs']],
  [/^public\/common\/scout\//, ['tests/scout.test.mjs']],
  [/^public\/[^/]+\/js\/painter\.js$/, ['tests/ui/painter.test.mjs']],
  [/^public\/common\/(lang|i18n|suite)\.js$|^public\/(?:index|(?:quire|ledger|lectern)\/index)\.html$|^tools\/ledger\/build\/template\.html$/, ['tests/chinese-language.test.mjs']],
  [/^tools\/i18n\.mjs$|^tools\/ledger\/build\/template\.html$/, ['tests/i18n.test.mjs']],
  [/^public\/common\/core\.js$|^public\/ledger\/js\/model\.js$/, ['tests/fonts.test.mjs']],
  [/^public\/common\/locales\.js$|^public\/quire\/js\/(app|dmodel|docx-write)\.js$|^public\/ledger\/js\/(app|xlsx-write)\.js$|^public\/lectern\/js\/(model|pptx-write)\.js$/, ['tests/locales.test.mjs']],
  [/^tests\/lib\/gestures\.mjs$/, ['tests/ui/gestures.test.mjs']],
  [/^tests\/lib\/input\.mjs$|^tools\/ooxml\/cdp\.mjs$/, ['tests/ui/smoke.test.mjs', 'tests/hostile/helpers.test.mjs']],
  [/^public\/[^/]+\.html$|^public\/[^/]+\/index\.html$|^tools\/csp\.mjs$|^tools\/ledger\/build\/|^tools\/oss-licenses\.py$/, ['tests/csp.test.mjs']],   // the pages' policy
  [/^public\/oss-licenses\.html$|^tools\/oss-licenses\.py$|^desktop\/(Cargo|patches|licenses)/, ['tests/licenses.test.mjs']],
  [/^desktop\/src\//, ['tests/desktop/host.test.mjs']],
  [/^public\/common\/core\.js$|^public\/quire\/js\/htmlio\.js$|^public\/ledger\/js\/(csv|clipboard)\.js$|^public\/lectern\/js\/app\.js$/, ['tests/hostile/export-html.test.mjs']],   // HTML we write: no markup from the document
  [/^tests\/.*\.test\.(mjs|js)$/, (f) => [f]],   // a test changed: itself
];
const picked = new Set();
for (const f of changed) for (const [re, t] of RULES) if (re.test(f)) for (const x of typeof t === 'function' ? t(f) : t) picked.add(x);
const isUI = (t) => t.startsWith('tests/ui/');
/** 'file :: pattern' runs only the tests whose names match pattern */
const split = (t) => { const [file, pattern] = t.split(' :: '); return pattern ? [[file], ['--test-name-pattern', pattern]] : [[file], []]; };
const order = [...picked].sort((a, b) => isUI(a) - isUI(b) || (SECONDS[a] || 1) - (SECONDS[b] || 1));
const total = order.reduce((s, t) => s + (SECONDS[t] || 1), 0);

console.log(`Changed (${what}): ${changed.length} file${changed.length === 1 ? '' : 's'}`);
if (!order.length) { console.log('No tests cover these files (docs, plans, tools without tests): nothing to run.'); process.exit(0); }
console.log(`Runs, about ${total < 90 ? total + ' s' : Math.round(total / 60) + ' min'}:\n${order.map((t) => '  ' + t).join('\n')}`);
if (dry) process.exit(0);

/* run: the Node tests together, then the UI tests one file at a time (they share the browser's focus) */
const results = [];
const run = (label, files, extra = []) => {
  const started = Date.now();
  const r = spawnSync(process.execPath, ['--test', ...extra, ...files], { cwd: root, encoding: 'utf8', shell: false });
  const out = r.stdout + r.stderr, n = (k) => +((out.match(new RegExp(`^ℹ ${k} (\\d+)`, 'm')) || [])[1] || 0);
  results.push({ label, ok: r.status === 0, pass: n('pass'), fail: n('fail'), todo: n('todo'), s: Math.round((Date.now() - started) / 1000) });
  const last = results[results.length - 1];
  console.log(`${last.ok ? '✔' : '✖'} ${label}: ${last.pass} passed${last.fail ? `, ${last.fail} failed` : ''}${last.todo ? `, ${last.todo} todo` : ''} (${last.s} s)`);
  if (!last.ok) console.log(out.slice(out.indexOf('✖ failing tests:') >>> 0).split('\n').slice(0, 30).join('\n'));
};
const expand = (t) => t.includes('*') ? fs.globSync(t, { cwd: root }).sort() : [t];
const node = order.filter((t) => !isUI(t)), ui = order.filter(isUI);
if (node.length) run('Node tests', [...new Set(node.flatMap(expand))]);
for (const t of ui) { const [files, extra] = split(t); run(t, files, ['--test-concurrency=1', ...extra]); }
const bad = results.filter((r) => !r.ok);
console.log(bad.length ? `\n${bad.length} of ${results.length} failed.` : `\nAll ${results.length} passed. Before a release: node tools/ui-full.mjs (the whole of tests/ui/).`);
process.exit(bad.length ? 1 : 0);
