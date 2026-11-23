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
  const bar = () => document.querySelector('button.cast-bar');
  function checkBar(frozen) {
    const el = bar(), label = 'Casting to TV Now' + (frozen ? ' · Picture frozen' : '');
    if (el?.textContent !== label || !el.closest('[role=status], [aria-live=polite]')) throw new Error('Expected ' + label);
    const r = el.getBoundingClientRect();
    if (r.top !== 0 || r.left !== 0 || r.right !== innerWidth || r.height < 1 || r.height > 25 ||
        document.elementFromPoint(10, r.height / 2) !== el) throw new Error('Cast bar is not visible across the top of the window');
    if (document.querySelector('.show')) {
      if (!el.closest('.show')) throw new Error('Cast bar did not follow the slide show');
    } else if (r.bottom > document.querySelector('.cmdbars').getBoundingClientRect().top) throw new Error('Cast bar covers the menus');
  }
  const color = async () => {
    const c = document.createElement('canvas'); c.width = 160; c.height = 90;
    const ctx = c.getContext('2d');
    let failure;
    const decoder = new VideoDecoder({ output(frame) { ctx.drawImage(frame, 0, 0, c.width, c.height); frame.close(); }, error(e) { failure = e; } });
    try {
      decoder.configure({ codec: 'avc1.42E01F' });
      decoder.decode(new EncodedVideoChunk({ type: 'key', timestamp: 0, data: stubCast.frames.at(-1) }));
      await decoder.flush();
      if (failure) throw failure;
      // Sample across the top half, including letterboxing. The sender's red strip
      // spans the frame at these points even where its white lettering interrupts it.
      const red = getComputedStyle(bar()).backgroundColor.match(/\d+/g).slice(0, 3).map(Number);
      const pixels = ctx.getImageData(0, 0, c.width, c.height).data;
      for (let y = 0; y < c.height / 2; y++) {
        if ([40, 55, 105, 120].every(x => red.every((v, i) => Math.abs(v - pixels[4 * (c.width * y + x) + i]) < 18)))
          throw new Error('The decoded TV picture contains the sender’s cast bar');
      }
      const center = 4 * (c.width * (c.height / 2) + c.width / 2);
      return Array.from(pixels.slice(center, center + 3));
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
  function freeze(value) {
    if (app === 'lectern' && L.show.active) {
      document.querySelector('.show [aria-label="Slide show menu"]').click();
      const item = [...document.querySelectorAll('.menu .m-item')].find(el => el._it.cmd === 'castTV');
      if (!item) throw new Error('Cast is missing from the slide show menu');
      item.click();
    } else bar().click();
    const toggle = document.querySelector('#cast-freeze');
    if (!toggle || toggle.checked === value) throw new Error('Unexpected Freeze picture state');
    toggle.click();
    [...document.querySelectorAll('.dlg-foot button')].find(b => b.textContent === 'Close').click();
    checkBar(value);
  }
  await wait(async () => window.L?.ui?.cmds.castTV && await window.VO?.castView?.());
  if (bar()) throw new Error('Cast bar present before casting');
  L.ui.exec('castTV');
  await wait(() => document.querySelector('#cast-tv')?.options.length);
  if (document.querySelectorAll('.dlg select').length !== 1) throw new Error('Expected only the TV selector');
  document.querySelector('.dlg .primary').click();
  await wait(() => stubCast.plays === 1 && !document.querySelector('.dlg'));
  checkBar(false);
  // A private overlay covers the sender. The TV must still see the document underneath it.
  const privateUI = cover(document.body, '#ff00ff');
  const documentLayer = () => app === 'quire' ? L.layout.root.querySelector('.pg') :
    app === 'ledger' ? document.querySelector('#glayer') : L.ed.host.firstElementChild;
  cover(documentLayer(), '#00ff00');
  const colors = [await sample()];
  // Suspend a render before it reads the document. An edit made after Freeze
  // must stay private even when that in-flight render eventually finishes.
  const view = VO.castView;
  let resume;
  VO.castView = async () => {
    VO.castView = view;
    await new Promise(resolve => { resume = resolve; });
    return view();
  };
  await wait(() => resume);
  freeze(true);
  cover(documentLayer(), '#ffff00');
  const beforeFreeze = stubCast.frames.at(-1);
  await wait(() => stubCast.frames.at(-1) !== beforeFreeze);
  resume();
  colors.push(await sample());
  if (app === 'lectern') {
    privateUI.remove();
    key('F5');
    await wait(() => L.show.active);
    checkBar(true);
    cover(document.querySelector('.show .sl'), '#ff0000');
    colors.push(await sample());
    freeze(false);
    colors.push(await sample());
    key('b');
    await wait(() => document.querySelector('.show .blank'));
    colors.push(await sample());
    key('Escape');
    await wait(() => !L.show.active);
    checkBar(false);
    cover(documentLayer(), '#0000ff');
    colors.push(await sample());
  } else {
    freeze(false);
    colors.push(await sample());
  }
  privateUI.remove();
  bar().click();
  document.querySelector('.dlg .primary').click();
  await wait(() => stubCast.stops === 1);
  if (bar()) throw new Error('Cast bar remains after Stop');
  const result = { colors, plays: stubCast.plays, stops: stubCast.stops, key: stubCast.key };
  const expected = app === 'lectern' ? [[0,255,0], [0,255,0], [0,255,0], [255,0,0], [0,0,0], [0,0,255]] : [[0,255,0], [0,255,0], [255,255,0]];
  if (result.plays !== 1 || result.stops !== 1 || result.key !== null ||
      colors.length !== expected.length || expected.some((rgb,i) => rgb.some((c,j) => Math.abs(c - colors[i][j]) >= 24)))
    throw new Error(app + ' unexpected cast output ' + JSON.stringify(result));
  await VO.settings.kept();
  if (VO.settings.get('vibeoffice.castTV') !== 'tv-1') throw new Error('Last TV was not saved');
  return app;
}
