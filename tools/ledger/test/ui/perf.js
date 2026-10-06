/* Time File > Open, scrolling and saving for large files in Chromium. Usage: node perf.js file1 [file2 ...] */
const { chromium } = require('playwright');
(async () => {
  const browser = await chromium.launch();
  const page = await (await browser.newContext({ viewport: { width: 1280, height: 820 }, acceptDownloads: true })).newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('http://127.0.0.1:8760/ledger/');
  await page.waitForTimeout(700);
  for (const f of process.argv.slice(2)) {
    const t0 = Date.now();
    const [fc] = await Promise.all([page.waitForEvent('filechooser'), page.evaluate(() => { L.app.openDialog(); })]);
    await fc.setFiles(f);
    const name = f.split('/').pop();
    await page.waitForFunction((n) => L.app.books[L.app.cur] && L.app.books[L.app.cur].name === n, name, { timeout: 300000, polling: 100 });
    const openMs = Date.now() - t0;
    const info = await page.evaluate(() => { const w = L.grid.wb; let cells = 0; for (const s of w.sheets) s.rows.forEach((r) => { if (r) for (const k in r.cells) cells++; }); return { sheets: w.sheets.length, cells, heapMB: performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1048576) : null }; });
    /* scroll: 40 page-downs, time per frame */
    const t1 = Date.now();
    for (let i = 0; i < 40; i++) await page.keyboard.press('PageDown');
    await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => r())));
    const scrollMs = (Date.now() - t1) / 40;
    /* save as xlsx */
    const t2 = Date.now();
    const bytes = await page.evaluate(async () => { const b = await L.xlsxWrite.bytes(L.grid.wb, { type: 'xlsx' }); return b.length; });
    const saveMs = Date.now() - t2;
    console.log(JSON.stringify({ file: name, openMs, ...info, scrollMsPerPage: Math.round(scrollMs), saveMs, savedBytes: bytes }));
  }
  console.log('errors:', errors.length ? errors : 'none');
  await browser.close();
})();
