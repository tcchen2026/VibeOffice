// Decode the actual WebCodecs output from each native app's Cast command.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const run = promisify(execFile);
const pre = 'tests/desktop/host-stub.js,DLNAnything/web/canvas.js';
async function shot(app, steps, query = 'new&cast', scripts = pre) {
  const { stdout } = await run(process.execPath, ['tools/shot.mjs', app + '/', JSON.stringify(steps),
    '--qs', query, '--pre', scripts]);
  assert.match(stdout, /no errors, no external requests/);
  return JSON.parse(stdout.split('\n').findLast(l => l.startsWith('js -> ')).slice(6));
}

async function exercise() {
  const app = document.documentElement.dataset.app;
  const wait = async (predicate) => {
    const start = performance.now();
    while (!await predicate()) {
      if (performance.now() - start > 8000) throw new Error(document.querySelector('.dlg-body')?.textContent || 'Cast timed out');
      await new Promise(r => setTimeout(r, 40));
    }
  };
  const key = (value) => document.activeElement.dispatchEvent(new KeyboardEvent('keydown', {
    key: value, code: value, bubbles: true, cancelable: true
  }));
  const cover = (parent, color) => {
    const el = document.createElement('div');
    Object.assign(el.style, { position: 'absolute', inset: '0', zIndex: '9999', background: color, pointerEvents: 'none' });
    parent.appendChild(el);
    return el;
  };
  const color = async () => {
    const c = document.createElement('canvas'); c.width = c.height = 4;
    const ctx = c.getContext('2d');
    let failure;
    const decoder = new VideoDecoder({ output(frame) { ctx.drawImage(frame, 0, 0, 4, 4); frame.close(); }, error(e) { failure = e; } });
    try {
      decoder.configure({ codec: 'avc1.42E01F' });
      decoder.decode(new EncodedVideoChunk({ type: 'key', timestamp: 0, data: stubCast.frames.at(-1) }));
      await decoder.flush();
      if (failure) throw failure;
      return Array.from(ctx.getImageData(2, 2, 1, 1).data).slice(0, 3);
    } finally { if (decoder.state !== 'closed') decoder.close(); }
  };
  async function sample() {
    const previous = stubCast.frames.at(-1);
    await wait(() => stubCast.frames.at(-1) !== previous);
    return color();
  }
  L.ui.exec('castTV');
  await wait(() => document.querySelector('#cast-tv')?.options.length);
  if (app === 'ledger') document.querySelector('#cast-mode').value = 'ts';
  document.querySelector('.dlg .primary').click();
  await wait(() => stubCast.plays === 1 && !document.querySelector('#cast-tv'));
  // A private overlay covers the sender. The TV must still see the document underneath it.
  const privateUI = cover(document.body, '#ff00ff');
  const documentLayer = () => app === 'quire' ? L.layout.root.querySelector('.pg') :
    app === 'ledger' ? document.querySelector('#glayer') : L.ed.host.firstElementChild;
  cover(documentLayer(), '#00ff00');
  const colors = [await sample()];
  if (app === 'lectern') {
    privateUI.remove();
    key('F5');
    await wait(() => L.show.active);
    cover(document.querySelector('.show .sl'), '#ff0000');
    colors.push(await sample());
    key('b');
    await wait(() => document.querySelector('.show .blank'));
    colors.push(await sample());
    key('Escape');
    await wait(() => !L.show.active);
    cover(documentLayer(), '#0000ff');
    colors.push(await sample());
  }
  privateUI.remove();
  L.ui.exec('castTV');
  document.querySelector('.dlg .primary').click();
  await wait(() => stubCast.stops === 1);
  return { colors, mode: stubCast.mode, plays: stubCast.plays, stops: stubCast.stops, key: stubCast.key };
}

for (const app of ['quire', 'ledger', 'lectern']) {
  test(app + ' casts document pixels through one receiver session', async () => {
    const result = await shot(app, [{ js: 'return (' + exercise.toString() + ')()' }]);
    assert.deepEqual([result.mode, result.plays, result.stops, result.key], [app === 'ledger' ? 'ts' : 'hls', 1, 1, null]);
    const expected = app === 'lectern' ? [[0,255,0],[255,0,0],[0,0,0],[0,0,255]] : [[0,255,0]];
    assert.equal(result.colors.length, expected.length);
    expected.forEach((rgb,i) => rgb.forEach((c,j) => assert.ok(Math.abs(c-result.colors[i][j]) < 24, JSON.stringify(result.colors))));
  });
}

test('the web edition has no receiver command or native capture API', async () => {
  const result = await shot('lectern', [{ js: 'return [VO.host.cast, !!L.ui.cmds.castTV, typeof window.DLNAnything]' }],
    'new&web', 'tests/desktop/host-stub.js');
  assert.deepEqual(result, [null, false, 'undefined']);
});
