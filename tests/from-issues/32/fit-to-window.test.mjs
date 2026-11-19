import { test } from 'node:test';
import assert from 'node:assert/strict';
import { open } from '../../lib/input.mjs';

test('Issue 32: Fit to Window stays steady at fractional display scaling', async () => {
  const tab = await open('lectern');
  try {
    const result = await tab.evaluate(`(async () => {
      await L.app.loadSample();
      const { wrap, scroller, host } = L.ed;
      // The half-pixel pane height rounds up in clientHeight. Scale just this pane in the test tab.
      Object.assign(wrap.style, { position: 'fixed', top: '0', left: '0', zoom: '1.5', width: '1322px', height: '871.5px' });
      L.ed.setZoom(2);
      scroller.scrollTo(scroller.scrollWidth, scroller.scrollHeight);
      const panned = scroller.scrollLeft > 0 && scroller.scrollTop > 0;

      L.ui.exec('zoomFit');
      const frame = () => new Promise(requestAnimationFrame);
      await frame();
      await frame();
      const widths = new Set();
      for (let i = 0; i < 8; i++) {
        await frame();
        widths.add(host.getBoundingClientRect().width);
      }
      return { panned, widths: [...widths], overflowX: scroller.scrollWidth - scroller.clientWidth,
        overflowY: scroller.scrollHeight - scroller.clientHeight };
    })()`);
    assert.ok(result.panned, 'the enlarged slide can scroll on both axes');
    assert.equal(result.widths.length, 1, 'slide widths across idle frames: ' + result.widths.join(', '));
    assert.equal(result.overflowX, 0, 'the fitted slide fits horizontally');
    assert.equal(result.overflowY, 0, 'the fitted slide fits vertically');
    tab.clean();
  } finally { await tab.close(); }
});
