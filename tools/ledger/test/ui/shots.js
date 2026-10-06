/* Open the named workbooks in Chromium and screenshot each one (visual fidelity check).
 * Usage: node shots.js outDir file1 [file2 ...]  (paths to .xlsx / .csv / .xml files) */
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const [OUT, ...files] = process.argv.slice(2);
fs.mkdirSync(OUT, { recursive: true });
(async () => {
  const browser = await chromium.launch();
  const page = await (await browser.newContext({ viewport: { width: 1280, height: 820 } })).newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('http://127.0.0.1:8760/ledger/');
  await page.waitForTimeout(700);
  await page.evaluate(() => L.app.toggleTask(false));
  for (const f of files) {
    const [fc] = await Promise.all([page.waitForEvent('filechooser'), page.evaluate(() => { L.app.openDialog(); })]);
    await fc.setFiles(f);
    await page.waitForTimeout(1500);
    const sheet = process.env.SHEET;
    if (sheet) await page.evaluate((s) => { const w = L.grid.wb; const i = w.sheets.findIndex((x) => x.name === s); if (i >= 0) L.app.activateSheet ? L.app.activateSheet(i) : L.grid.showSheet(i); }, sheet);
    await page.waitForTimeout(300);
    const out = path.join(OUT, path.basename(f).replace(/\W+/g, '_').slice(0, 70) + '.png');
    await page.screenshot({ path: out });
    console.log(out);
  }
  console.log('errors:', errors.length ? errors : 'none');
  await browser.close();
})();
