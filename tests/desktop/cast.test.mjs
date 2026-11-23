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

test('TVs appear during discovery, the remembered TV is selected, and Cast can start early', async () => {
  const result = await shot('ledger', [{ js: until + `
    VO.settings.set('vibeoffice.castTV', 'tv-1');
    let next, calls = 0;
    const invoke = __TAURI_INTERNALS__.invoke;
    __TAURI_INTERNALS__.invoke = async (cmd, args, opts) => {
      if (cmd !== 'host_cast_scan') return invoke(cmd, args, opts);
      calls++;
      if (calls === 1) return { devices: [{id:'other', name:'Other TV'}], done:false };
      return new Promise(r => { next = r; });
    };
    L.ui.exec('castTV');
    const selector = () => document.querySelector('#cast-tv');
    await until(() => selector()?.length === 1 && next);
    const early = [selector().value, document.querySelector('.dlg .primary').disabled,
      document.querySelector('.dlg [role=status]').textContent.includes('Searching')];
    cancel(); await delay(0);
    L.ui.exec('castTV');
    await until(() => selector()?.length === 1);
    next({ devices: [{id:'tv-1', name:'Office TV'}], done:false });
    await until(() => calls === 3 && selector().length === 2);
    const preferred = selector().value;
    document.querySelector('.dlg .primary').click();
    await until(() => stubCast.plays === 1);
    next({devices:[], done:true});
    await until(() => !selector());
    return [early, preferred, calls, stubCast.plays];
  ` }]);
  assert.deepEqual(result, [['other', false, true], 'tv-1', 3, 1]);
});

test('waiting for the TV shows firewall help, keeps casting on dismissal and can stop the stream', async () => {
  const result = await shot('quire', [{ js: until + `
    let release;
    VO.host.cast.wait = () => new Promise(r => { release = r; });
    const status = () => document.querySelector('.dlg [role=status]')?.textContent || '';
    const button = label => [...document.querySelectorAll('.dlg-foot button')].find(b => b.textContent === label);
    async function start() {
      release = null;
      L.ui.exec('castTV');
      await until(() => document.querySelector('#cast-tv')?.length);
      button('Cast').click();
      await until(() => release && status().includes('Office TV'));
    }
    await start();
    const waiting = status();
    release(false);
    await until(() => document.querySelector('.dlg-body').innerText.includes('Windows Defender Firewall'));
    const help = document.querySelector('.dlg-body').innerText;
    button('Keep waiting').click();
    await until(() => !document.querySelector('.dlg'));
    const before = stubCast.frames.at(-1);
    await until(() => stubCast.frames.at(-1) !== before);
    const active = [stubCast.plays, stubCast.stops];
    release(true);
    L.ui.exec('castTV'); button('Stop casting').click();
    await until(() => stubCast.stops === 1);
    await start(); release(false);
    await until(() => document.querySelector('.dlg-body').innerText.includes('Windows Defender Firewall'));
    button('Stop casting').click();
    await until(() => stubCast.stops === 2);
    release(true);
    return [waiting, help.includes('same network'), help.includes('Private'), active, stubCast.plays, stubCast.stops];
  ` }]);
  assert.deepEqual(result, ['Waiting for Office TV...', true, true, [1, 0], 2, 2]);
});

test('the picker gives a plain playback error without displaying the receiver response', async () => {
  const message = await shot('ledger', [{ js: until + `
    VO.host.cast.play = async () => { throw {kind:'refused', detail:'<s:Envelope>internal fault</s:Envelope>'}; };
    L.ui.exec('castTV');
    await until(() => document.querySelector('#cast-tv')?.length);
    document.querySelector('.dlg .primary').click();
    await until(() => stubCast.stops && !document.querySelector('.dlg .primary').disabled);
    return document.querySelector('.dlg [role=status]').textContent;
  ` }]);
  assert.match(message, /TV declined to play/);
  assert.doesNotMatch(message, /Envelope|fault|object Object/);
});

test('a rejected picture stops the cast and identifies the picture failure', async () => {
  const result = await shot('ledger', [{ js: until + `
    const frame = VO.host.cast.frame;
    VO.host.cast.frame = (key, bytes) => {
      if (stubCast.plays) throw {kind:'picture', detail:'invalid Annex B'};
      return frame(key, bytes);
    };
    L.ui.exec('castTV');
    await until(() => document.querySelector('#cast-tv')?.length);
    document.querySelector('.dlg .primary').click();
    await until(() => stubCast.stops === 1 && document.querySelector('.dlg-body'));
    return document.querySelector('.dlg-body').innerText;
  ` }]);
  assert.match(result, /Casting stopped because the next picture could not be sent/);
  assert.doesNotMatch(result, /Annex B|local network/);
});
