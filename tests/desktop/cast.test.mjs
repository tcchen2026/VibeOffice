// Decode the actual WebCodecs output from each native app's Cast command.
import { test } from 'node:test';
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const run = promisify(execFile);
const pre = 'tests/desktop/host-stub.js,DLNAnything/web/canvas.js';
const exercise = readFileSync(new URL('./cast-exercise.js', import.meta.url), 'utf8');
async function shot(app, steps, query = 'new&cast', scripts = pre) {
  const { stdout } = await run(process.execPath, ['tools/shot.mjs', app + '/', JSON.stringify(steps),
    '--qs', query, '--pre', scripts]);
  assert.match(stdout, /no errors, no external requests/);
  assert.doesNotMatch(stdout, /^js error/m);
  const result = stdout.split('\n').findLast(l => l.startsWith('js -> '));
  assert.ok(result, stdout);
  return JSON.parse(result.slice(6));
}

for (const app of ['quire', 'ledger', 'lectern']) {
  test(app + ' casts document pixels through one receiver session', async () => {
    assert.equal(await shot(app, [{ js: exercise + '\nreturn exerciseCast()' }]), app);
  });
}

test('a real encoder restricted to AVC still casts an independently decodable picture', async () => {
  const setup = `
    const probe = VideoEncoder.isConfigSupported.bind(VideoEncoder);
    let usedAvc = false;
    VideoEncoder.isConfigSupported = config => {
      if (config.avc.format === 'annexb') return { supported: false };
      usedAvc = true;
      return probe(config);
    };
  `;
  assert.deepEqual(await shot('ledger', [{ js: setup + exercise + '\nreturn [await exerciseCast(), usedAvc]' }]), ['ledger', true]);
});

const mac = 'new&cast&appstore&lang=en';
const until = `
  const delay = ms => new Promise(r => setTimeout(r, ms));
  async function until(ok) {
    const start = performance.now();
    while (!ok()) {
      if (performance.now() - start > 5000) throw new Error('Timed out: ' + document.querySelector('.dlg-body')?.innerText);
      await delay(20);
    }
  }
  const cancel = () => [...document.querySelectorAll('.dlg-foot button')].find(b => b.textContent === 'Cancel').click();
`;

test('macOS checks H.264 before any discovery, even if the user cancels during the check', async () => {
  const result = await shot('quire', [{ js: until + `
    if (stubSession.hostCalls?.includes('host_cast_')) throw new Error('Cast accessed the host on launch');
    let scans = 0, finish;
    VO.host.cast.scan = async () => { scans++; return []; };
    window.DLNAnything = { isSupported: async () => false };
    L.ui.exec('castTV');
    await until(() => !document.querySelectorAll('.dlg-foot button')[1].disabled);
    const message = document.querySelector('.dlg [role=status]').textContent;
    cancel(); await delay(0);
    window.DLNAnything.isSupported = () => new Promise(r => { finish = r; });
    L.ui.exec('castTV');
    cancel();
    finish(true); await delay(0);
    return [message, scans];
  ` }], mac);
  assert.deepEqual(result, ['This webview cannot encode H.264 for casting.', 0]);
});

test('macOS retries discovery after a permission prompt and cancels pending retries on close', async () => {
  const result = await shot('lectern', [{ js: until + `
    const findTV = VO.host.cast.scan;
    let scans = 0, finish;
    VO.host.cast.scan = async () => {
      scans++;
      if (scans === 1) return new Promise((_, reject) => { finish = reject; });
      if (scans === 2) return findTV();
      return [];
    };
    L.ui.exec('castTV');
    await until(() => finish);
    const prompt = document.querySelector('.dlg [role=status]').textContent;
    finish(new Error('Network permission denied'));
    await until(() => document.querySelector('.dlg-body')?.innerText.includes('System Settings'));
    const help = document.querySelector('.dlg-body').innerText;
    window.dispatchEvent(new Event('focus'));
    window.dispatchEvent(new Event('focus'));
    await until(() => document.querySelector('#cast-tv')?.length);
    const found = document.querySelector('#cast-tv').selectedOptions[0].text;
    cancel(); await delay(0);
    L.ui.exec('castTV');
    await until(() => scans === 3 && document.querySelector('.dlg-body').innerText.includes('System Settings'));
    cancel(); window.dispatchEvent(new Event('focus'));
    await delay(1650);
    return [prompt, help.includes('Privacy & Security → Local Network'), found, scans, stubCast.plays];
  ` }], mac);
  assert.match(result[0], /Local Network.*Allow/);
  assert.deepEqual(result.slice(1), [true, 'Office TV', 3, 0]);
});

test('macOS makes one timed retry when its permission sheet leaves the window focused', async () => {
  const result = await shot('ledger', [{ js: until + `
    let scans = 0;
    VO.host.cast.scan = async () => { scans++; return []; };
    L.ui.exec('castTV');
    await until(() => scans === 2);
    await delay(1650);
    cancel();
    return scans;
  ` }], mac);
  assert.equal(result, 2);
});

test('reopening the macOS picker during discovery shares the outstanding search', async () => {
  const result = await shot('quire', [{ js: until + `
    let scans = 0, reply;
    VO.host.cast.scan = () => { scans++; return new Promise(r => { reply = r; }); };
    L.ui.exec('castTV');
    await until(() => reply);
    cancel(); await delay(0);
    L.ui.exec('castTV');
    await until(() => document.querySelector('.dlg [role=status]').textContent.includes('Searching'));
    reply([{ id: 'tv-1', name: 'Office TV' }]);
    await until(() => document.querySelector('#cast-tv')?.length);
    cancel();
    return [scans, stubCast.plays];
  ` }], mac);
  assert.deepEqual(result, [1, 0]);
});

test('the web edition has no receiver command or native capture API', async () => {
  const result = await shot('lectern', [{ js: 'return [VO.host.cast, !!L.ui.cmds.castTV, typeof window.DLNAnything]' }],
    'new&web', 'tests/desktop/host-stub.js');
  assert.deepEqual(result, [null, false, 'undefined']);
});
