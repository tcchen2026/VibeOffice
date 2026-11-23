// Run inside either Chromium or macOS WKWebView with the native host fixture.
async function exerciseCast() {
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
    // A frame can already be in flight when the document changes.
    for (let i = 0; i < 2; i++) {
      const previous = stubCast.frames.at(-1);
      await wait(() => stubCast.frames.at(-1) !== previous);
    }
    return color();
  }
  await wait(async () => window.L?.ui?.cmds.castTV && await window.VO?.castView?.());
  L.ui.exec('castTV');
  await wait(() => document.querySelector('#cast-tv')?.options.length);
  if (document.querySelectorAll('.dlg select').length !== 1) throw new Error('Expected only the TV selector');
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
  const result = { colors, plays: stubCast.plays, stops: stubCast.stops, key: stubCast.key };
  const expected = app === 'lectern' ? [[0,255,0], [255,0,0], [0,0,0], [0,0,255]] : [[0,255,0]];
  if (result.plays !== 1 || result.stops !== 1 || result.key !== null ||
      colors.length !== expected.length || expected.some((rgb,i) => rgb.some((c,j) => Math.abs(c - colors[i][j]) >= 24)))
    throw new Error(app + ' unexpected cast output ' + JSON.stringify(result));
  await VO.settings.kept();
  if (VO.settings.get('vibeoffice.castTV') !== 'tv-1') throw new Error('Last TV was not saved');
  return app;
}
