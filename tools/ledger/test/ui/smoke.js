/* Load Ledger in Chromium, report console errors, take a screenshot. Usage: node smoke.js [out.png] */
const { chromium } = require('playwright');
(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  const errors = [];
  page.on('console', (m) => { if (m.type() === 'error' || m.type() === 'warning') errors.push(m.type() + ': ' + m.text()); });
  page.on('pageerror', (e) => errors.push('pageerror: ' + e.message + '\n' + (e.stack || '').split('\n').slice(0, 12).join('\n')));
  await page.goto('http://127.0.0.1:8760/ledger/');
  await page.waitForTimeout(1500);
  await page.screenshot({ path: process.argv[2] || '/tmp/shots/smoke.png' });
  console.log(errors.join('\n') || 'no errors');
  await browser.close();
})();
