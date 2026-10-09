// Mouse-wheel and keyboard scrolling in every direction, with the page repainting between steps
// (the scroll clamp once used the pane's current first row, so up/left stopped where they were).
// node tools/ledger/test/scroll-ui.mjs
import { openPage, ensureServer } from '../../ooxml/cdp.mjs';
const stop = await ensureServer(), page = await openPage('ledger', { persistence: false });
try {
  const result = await page.evaluate(`(${run.toString()})()`);
  if (page.errors.length) result.errors = page.errors;
  console.log(JSON.stringify(result));
  if (result.failures.length || page.errors.length) process.exitCode = 1;
} finally { await page.close(); stop(); }

async function run() {
  const G = L.grid, cv = document.querySelector('canvas'), b = cv.getBoundingClientRect(), failures = [];
  const pause = () => new Promise(r => setTimeout(r, 60));
  const at = () => [G.vs().scrollR, G.vs().scrollC];
  const wheel = async (dx, dy, times) => {
    for (let i = 0; i < times; i++) {
      cv.dispatchEvent(new WheelEvent('wheel', { deltaX: dx, deltaY: dy, deltaMode: 1, bubbles: true, cancelable: true, clientX: b.x + 300, clientY: b.y + 300 }));
      await pause();
    }
  };
  const expect = (what, actual, wanted) => { if (JSON.stringify(actual) !== JSON.stringify(wanted)) failures.push({ what, actual, wanted }); };
  async function directions(label, home) {
    G.vs().scrollR = home[0]; G.vs().scrollC = home[1]; G.paint(true); await pause();
    await wheel(0, 3, 5); expect(label + ' down', at(), [home[0] + 15, home[1]]);
    await wheel(0, -3, 5); expect(label + ' up', at(), home);
    await wheel(3, 0, 5); expect(label + ' right', at(), [home[0], home[1] + 15]);
    await wheel(-3, 0, 5); expect(label + ' left', at(), home);
    await wheel(0, -3, 2); expect(label + ' stops at the first scrollable row', at(), home);
  }
  await directions('plain', [0, 0]);
  // Arrow keys above the view scroll it up with the selection.
  await wheel(0, 3, 4); G.select(G.vs().scrollR + 2, 0); await pause();
  const key = k => document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }));
  for (let i = 0; i < 6; i++) { key('ArrowUp'); await pause(); }
  expect('arrow up reveals the selection', G.vs().scrollR <= G.sel().r, true);
  // Frozen panes: two rows and one column; scrolling stops below and right of them.
  G.sheet().view.freeze = { r: 2, c: 1 }; G.sheet().view.top = { r: 0, c: 0 };
  await directions('frozen', [2, 1]);
  return { failures };
}
