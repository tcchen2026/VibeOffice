/* Encrypted workbooks in the browser: open with the right password, refuse a wrong one, save with a
 * password and reopen. Usage: node password.js encrypted.xlsx password */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const [file, pw] = process.argv.slice(2);
(async () => {
  const browser = await chromium.launch();
  const page = await (await browser.newContext({ viewport: { width: 1280, height: 820 }, acceptDownloads: true })).newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto('http://127.0.0.1:8760/ledger/');
  await page.waitForTimeout(700);
  let pass = 0, fail = 0;
  const ok = (c, n, info) => { if (c) { pass++; console.log('PASS', n); } else { fail++; console.log('FAIL', n, info || ''); } };
  const open = async (p, password) => {
    const [fc] = await Promise.all([page.waitForEvent('filechooser'), page.evaluate(() => { L.app.openDialog(); })]);
    await fc.setFiles(p);
    const box = await page.waitForSelector('.dlg input[type=password]', { timeout: 15000 });
    await box.fill(password);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(500);
  };
  /* wrong password: message, nothing opened */
  const before = await page.evaluate(() => L.app.books.length);
  await open(file, 'not-the-password');
  await page.waitForSelector('.dlg', { timeout: 20000 });
  const msg = await page.evaluate(() => { const d = Array.from(document.querySelectorAll('.dlg')).pop(); return d ? d.textContent : ''; });
  ok(/not correct/i.test(msg), 'wrong password refused', msg.slice(0, 120));
  await page.keyboard.press('Enter');
  ok((await page.evaluate(() => L.app.books.length)) === before, 'nothing opened after wrong password');
  /* right password */
  await open(file, pw);
  await page.waitForFunction((n) => L.app.books[L.app.cur].name === n, path.basename(file), { timeout: 30000 });
  ok(true, 'opened with the right password');
  /* save with a new password, reopen */
  const bytes = await page.evaluate(async () => Array.from(await L.xlsxWrite.bytes(L.grid.wb, { type: 'xlsx', password: 'n3w-Pass' })));
  const out = path.join(path.dirname(file), 'pw-resaved.xlsx');
  fs.writeFileSync(out, Buffer.from(bytes));
  await open(out, 'n3w-Pass');
  await page.waitForFunction(() => L.app.books[L.app.cur].name === 'pw-resaved.xlsx', null, { timeout: 30000 });
  ok(true, 'saved with a password and reopened');
  console.log(`${pass} passed, ${fail} failed; errors: ${errors.length ? errors.join(' | ') : 'none'}`);
  await browser.close();
})();
