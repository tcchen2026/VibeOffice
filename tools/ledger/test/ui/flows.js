/* Scripted end-to-end flows in Chromium: typing, formulas, formatting, open/save xlsx and csv, chart, filter,
 * print preview, dialogs. Prints PASS/FAIL lines and writes screenshots to the output directory. */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const OUT = process.argv[2] || '/tmp/shots/ui';
const CORPUS = path.join(process.env.VO_CORPORA || path.join(require('os').homedir(), 'corpora'), 'excel');   // tools/corpora.sh
fs.mkdirSync(OUT, { recursive: true });
let pass = 0, fail = 0;
const ok = (cond, name, info) => { process.stderr.write('.'); if (cond) { pass++; console.log('PASS', name); } else { fail++; console.log('FAIL', name, info === undefined ? '' : JSON.stringify(info).slice(0, 300)); } };
(async () => {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 820 }, acceptDownloads: true });
  const page = await ctx.newPage();
  page.setDefaultTimeout(8000);
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message + ' | ' + (e.stack || '').split('\n').slice(1, 3).join(' / ')));
  page.on('console', (m) => { if (m.type() === 'error' && !/ERR_TUNNEL|Failed to load resource/.test(m.text())) errors.push('console: ' + m.text().slice(0, 300)); });
  await page.goto('http://127.0.0.1:8760/ledger/');
  await page.waitForTimeout(700);
  const ev = (fn, arg) => Promise.race([page.evaluate(fn, arg), new Promise((res, rej) => setTimeout(() => rej(new Error('TIMEOUT in ' + String(fn).slice(0, 120))), 15000))]);
  const shot = (name) => page.screenshot({ path: path.join(OUT, name + '.png') });
  const goto = async (ref) => { if (process.env.V) console.log('goto', ref); await page.click('#namebox'); await page.fill('#namebox', ref); await page.keyboard.press('Enter'); await page.waitForTimeout(50); };
  const val = (ref, sheet) => ev(([ref, sheet]) => { const w = L.grid.wb; const s = sheet == null ? L.grid.sheet() : w.sheets[sheet]; const p = L.formula.parseCell(ref); const v = L.calc.cellValue(s, p.r, p.c); return v && v.e ? v.e : v; }, [ref, sheet]);

  /* ---------- 1. typing values and formulas ---------- */
  await ev(() => L.app.newWorkbook());
  await page.waitForTimeout(100);
  await goto('A1');
  await page.keyboard.type('Item'); await page.keyboard.press('Tab');
  await page.keyboard.type('Qty'); await page.keyboard.press('Tab');
  await page.keyboard.type('Price'); await page.keyboard.press('Tab');
  await page.keyboard.type('Total'); await page.keyboard.press('Enter');
  const rows = [['Apples', '12', '0.5'], ['Pears', '7', '0.75'], ['Plums', '30', '0.2'], ['Kiwis', '5', '1.1']];
  for (let i = 0; i < rows.length; i++) {
    await goto('A' + (i + 2));
    for (const x of rows[i]) { await page.keyboard.type(x); await page.keyboard.press('Tab'); }
    await page.keyboard.type(`=B${i + 2}*C${i + 2}`); await page.keyboard.press('Enter');
  }
  await goto('D6'); await page.keyboard.type('=SUM(D2:D5)'); await page.keyboard.press('Enter');
  ok(await val('D2') === 6, 'formula D2 = B2*C2', await val('D2'));
  ok(Math.abs(await val('D6') - (6 + 5.25 + 6 + 5.5)) < 1e-9, 'SUM over typed range', await val('D6'));
  await goto('B3'); await page.keyboard.type('10'); await page.keyboard.press('Enter');
  ok(Math.abs(await val('D6') - (6 + 7.5 + 6 + 5.5)) < 1e-9, 'recalc after edit', await val('D6'));
  /* typed date, percent, currency, text with apostrophe */
  await goto('F1'); await page.keyboard.type('3/14/2026'); await page.keyboard.press('Enter');
  await page.keyboard.type('12%'); await page.keyboard.press('Enter');
  await page.keyboard.type('$1,250.50'); await page.keyboard.press('Enter');
  await page.keyboard.type("'00123"); await page.keyboard.press('Enter');
  const fmt = await ev(() => { const s = L.grid.sheet(); return [0, 1, 2, 3].map((r) => [s.val(r, 5), s.wb.styles.get((s.get(r, 5) || {}).s || 0).nf]); });
  ok(fmt[0][0] === 46095 && /m\/d\/yyyy/.test(fmt[0][1]), 'date entry', fmt[0]);
  ok(Math.abs(fmt[1][0] - 0.12) < 1e-12 && fmt[1][1] === '0%', 'percent entry', fmt[1]);
  ok(fmt[2][0] === 1250.5 && /\$/.test(fmt[2][1]), 'currency entry', fmt[2]);
  ok(fmt[3][0] === '00123', 'apostrophe text', fmt[3]);
  /* undo / redo */
  await page.keyboard.press('Control+z');
  ok(await ev(() => L.grid.sheet().val(3, 5)) == null, 'undo removes last entry');
  await page.keyboard.press('Control+y');
  ok(await ev(() => L.grid.sheet().val(3, 5)) === '00123', 'redo restores entry');
  /* bold via shortcut, number format via command */
  await goto('A1:D1'); await page.keyboard.press('Control+b');
  ok(await ev(() => !!L.grid.wb.styles.get(L.grid.sheet().get(0, 2).s).font.b), 'Ctrl+B bolds header');
  await goto('C2:D6'); await ev(() => L.ui.exec('currencyStyle'));
  ok(await ev(() => /\$/.test(L.grid.wb.styles.get(L.grid.sheet().get(2, 3).s).nf)), 'Currency Style applied');
  /* AutoSum button on a selection */
  await goto('B2:B5'); await ev(() => L.app.autoSum('SUM'));
  ok(await val('B6') === 57, 'AutoSum below selection', await val('B6'));
  /* fill handle equivalent: AutoFill series */
  await goto('H1'); await page.keyboard.type('Mon'); await page.keyboard.press('Enter');
  await ev(() => { const s = L.grid.sheet(); L.ops.autoFill(s, { r1: 0, c1: 7, r2: 0, c2: 7 }, { r1: 0, c1: 7, r2: 6, c2: 7 }, 'auto'); });
  ok(await ev(() => L.grid.sheet().val(6, 7)) === 'Sun', 'AutoFill weekdays', await ev(() => L.grid.sheet().val(6, 7)));
  /* insert a row above 3 and check references */
  await goto('A3'); await ev(() => L.app.insertLines('r'));
  ok(await ev(() => L.grid.sheet().get(6, 3).f) === 'SUM(D2:D6)', 'insert row adjusts SUM', await ev(() => L.grid.sheet().get(6, 3).f));
  await ev(() => L.app.undo());
  ok(await ev(() => L.grid.sheet().get(5, 3).f) === 'SUM(D2:D5)', 'undo insert row');
  await shot('01-typing');

  /* ---------- 2. dialogs render ---------- */
  await goto('C3'); await ev(() => L.ui.exec('formatCells'));
  await page.waitForTimeout(150); await shot('02-format-cells');
  await page.click('.tabs-head .tab:nth-child(4)'); await page.waitForTimeout(80); await shot('02b-format-border');
  await page.keyboard.press('Escape');
  await ev(() => L.ui.exec('pageSetup')); await page.waitForTimeout(100); await shot('03-page-setup'); await page.keyboard.press('Escape');
  await ev(() => L.ui.exec('insertFunction')); await page.waitForTimeout(100); await shot('04-insert-function'); await page.keyboard.press('Escape');
  await ev(() => L.ui.exec('sortDlg')); await page.waitForTimeout(100); await shot('05-sort'); await page.keyboard.press('Escape');
  await page.waitForTimeout(100);

  /* ---------- 3. sort, filter ---------- */
  await goto('A2'); await ev(() => L.ui.exec('sortAsc'));
  ok(await ev(() => L.grid.sheet().val(1, 0)) === 'Apples' && await ev(() => L.grid.sheet().val(4, 0)) === 'Plums', 'Sort ascending by Item', await ev(() => [1, 2, 3, 4].map((r) => L.grid.sheet().val(r, 0))));
  ok(await val('D2') === 6, 'formulas move with sorted rows', await val('D2'));
  await goto('A1'); await ev(() => L.ui.exec('autoFilter'));
  ok(await ev(() => !!L.grid.sheet().autoFilter), 'AutoFilter on');
  await page.waitForTimeout(100);
  const btns = await page.$$('.af-btn');
  ok(btns.length >= 4, 'filter buttons drawn', btns.length);
  if (btns.length) { await btns[0].click(); await page.waitForTimeout(100); await shot('06-filter-list'); const item = await page.$('.af-item:text-is("Pears")'); if (item) await item.click(); await page.waitForTimeout(100); }
  const hidden = await ev(() => L.grid.sheet().rows.filter((r) => r && r.hidden).length);
  ok(hidden === 4, 'filter hides 4 rows (Pears only, incl. total row)', hidden);
  await shot('07-filtered');
  await ev(() => L.ui.exec('showAll'));
  ok(await ev(() => L.grid.sheet().rows.filter((r) => r && r.hidden).length) === 0, 'Show All');
  await ev(() => L.ui.exec('autoFilter'));

  /* ---------- 4. chart ---------- */
  await goto('A1:B5');
  if (process.env.PAUSE) {
    const cdp = await ctx.newCDPSession(page);
    await cdp.send('Debugger.enable');
    cdp.on('Debugger.paused', (e) => { console.log('PAUSED'); for (const f of e.callFrames.slice(0, 15)) console.log('  at', f.functionName, f.url.split('/').pop(), f.location.lineNumber + 1, f.location.columnNumber); process.exit(0); });
    page.evaluate(() => L.chartWizard.open()).catch(() => {});
    await new Promise((r) => setTimeout(r, 3000));
    await cdp.send('Debugger.pause');
    await new Promise((r) => setTimeout(r, 5000));
  }
  await ev(() => { L.chartWizard.open(); });
  await page.waitForTimeout(200); await shot('08-chart-wizard');
  await page.click('.dlg-foot button:has-text("Finish")');
  await page.waitForTimeout(300);
  ok(await ev(() => L.grid.sheet().drawings.length === 1 && L.grid.sheet().drawings[0].kind === 'chart'), 'chart inserted');
  await shot('09-chart');
  await ev(() => L.grid.deselectObject());

  /* ---------- 5. save as xlsx and reopen ---------- */
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 15000 }), ev(() => L.app.saveAs('xlsx', 'flowtest.xlsx'))]);
  const saved = path.join(OUT, 'flowtest.xlsx');
  await dl.saveAs(saved);
  ok(fs.statSync(saved).size > 3000, 'saved xlsx', fs.statSync(saved).size);
  const [fc] = await Promise.all([page.waitForEvent('filechooser'), ev(() => { L.app.openDialog(); })]);
  await fc.setFiles(saved);
  await page.waitForTimeout(800);
  ok(await ev(() => L.app.books[L.app.cur].name) === 'flowtest.xlsx', 'reopened saved file', await ev(() => L.app.books[L.app.cur].name));
  ok(await val('B6') === 57 && await ev(() => L.grid.sheet().drawings.length) === 1, 'reopened values and chart', [await val('B6'), await ev(() => L.grid.sheet().drawings.length)]);
  await shot('10-reopened');

  /* ---------- 6. open real workbooks from the corpus ---------- */
  const samples = fs.readdirSync(CORPUS).filter((f) => /\.(xlsx|xlsm)$/.test(f) && fs.statSync(path.join(CORPUS, f)).size > 15000 && fs.statSync(path.join(CORPUS, f)).size < 400000);
  const pick = [];
  for (const k of ['chart', 'format', 'style', 'cond', 'merge', 'comment', 'table', 'image', 'freeze', 'border']) { const f = samples.find((x) => x.toLowerCase().includes(k) && !pick.includes(x)); if (f) pick.push(f); }
  for (const f of pick.slice(0, 10)) {
    const [fc2] = await Promise.all([page.waitForEvent('filechooser'), ev(() => { L.app.openDialog(); })]);
    await fc2.setFiles(path.join(CORPUS, f));
    await page.waitForTimeout(1200);
    const name = await ev(() => L.app.books[L.app.cur].name);
    ok(name === f, 'open ' + f, name);
    await shot('20-' + f.replace(/\W+/g, '_').slice(0, 60));
  }

  /* ---------- 7. CSV through the import wizard ---------- */
  const csvPath = path.join(OUT, 'sample.csv');
  fs.writeFileSync(csvPath, 'Name,Joined,Amount,Note\n"Smith, Jo",2024-01-15,"1,234.50",ok\nLee,3/2/2025,99,"multi\nline"\n');
  const [fc3] = await Promise.all([page.waitForEvent('filechooser'), ev(() => { L.app.openDialog(); })]);
  await fc3.setFiles(csvPath);
  await page.waitForTimeout(600);
  ok(await ev(() => L.grid.sheet().val(1, 0)) === 'Smith, Jo' && await ev(() => L.grid.sheet().val(1, 2)) === 1234.5, 'CSV opened with quoting and numbers', await ev(() => [L.grid.sheet().val(1, 0), L.grid.sheet().val(1, 2), L.grid.sheet().val(2, 3)]));
  await shot('30-csv');

  /* ---------- 8. print preview ---------- */
  await ev(() => L.app.switchBook(0));
  await ev(() => { L.print.preview(); }); await page.waitForTimeout(800); await shot('40-print-preview');
  ok(await page.$('.pp-page canvas') != null, 'print preview page drawn');
  await page.keyboard.press('Escape');

  /* ---------- 9. validation circles, page break preview drag, web page open ---------- */
  await ev(() => { L.app.newWorkbook(); });
  await page.waitForTimeout(300);
  await ev(() => {
    const sh = L.grid.sheet();
    L.ops.tx(sh.wb, 'data', () => { for (let r = 0; r < 120; r++) for (let c = 0; c < 4; c++) L.ops.setValue(sh, r, c, r + c); });
    L.ops.tx(sh.wb, 'dv', () => L.ops.setDV(sh, [{ type: 'whole', op: 'between', f1: '0', f2: '50', ranges: [{ r1: 0, c1: 0, r2: 119, c2: 0 }], showError: true }]));
  });
  await ev(() => L.ui.exec('circleInvalid'));
  await page.waitForTimeout(200);
  const circles = await ev(() => (L.grid.state.circles ? L.grid.state.circles.list.length : 0));
  ok(circles === 69, 'Circle Invalid Data marks the 69 values above 50', circles);
  await ev(() => L.ui.exec('clearCircles'));
  ok((await ev(() => L.grid.state.circles)) == null, 'Clear Validation Circles');
  await ev(() => L.app.setPageBreakPreview(true));
  await page.waitForTimeout(200);
  await page.keyboard.press('Enter');
  const brk = await ev(() => { const sh = L.grid.sheet(); const br = sh._pages.rows[0]; const fr = L.grid.frame(); const pane = fr.panes[fr.panes.length - 1]; const rect = document.querySelector('#gridwrap').getBoundingClientRect(); return { r: br.r, x: rect.left + L.render.colX(fr, pane, 1) + 8, y: rect.top + L.render.rowY(fr, pane, br.r), y2: rect.top + L.render.rowY(fr, pane, br.r - 10) }; });
  await page.mouse.move(brk.x, brk.y); await page.mouse.down(); await page.mouse.move(brk.x, brk.y - 20, { steps: 4 }); await page.mouse.move(brk.x, brk.y2, { steps: 4 }); await page.mouse.up();
  const manual = await ev(() => L.grid.sheet().print.rowBreaks);
  ok(manual.length === 1 && manual[0] === brk.r - 10, 'page break dragged in Page Break Preview', { manual, auto: brk.r });
  await ev(() => L.app.setPageBreakPreview(false));
  const htmlFile = path.join(OUT, 'page.html');
  fs.writeFileSync(htmlFile, '<html><body><h2>Prices</h2><table><tr><td>Tea</td><td>3.50</td></tr><tr><td>Coffee</td><td>$4.25</td></tr></table><h2>Stock</h2><table><tr><td>Tea</td><td>12</td></tr></table></body></html>');
  const [fc9] = await Promise.all([page.waitForEvent('filechooser'), ev(() => { L.app.openDialog(); })]);
  await fc9.setFiles(htmlFile);
  await page.waitForTimeout(1000);
  const web = await ev(() => ({ sheets: L.grid.wb.sheets.map((s) => s.name), v: L.grid.wb.sheets[0].val(1, 1), w: L.grid.wb.sheets[1] && L.grid.wb.sheets[1].val(0, 1) }));
  ok(web.sheets.join() === 'Prices,Stock' && web.v === 4.25 && web.w === 12, 'web page opened, one sheet per table', web);

  console.log('\nerrors:', errors.length ? '\n' + errors.join('\n') : 'none');
  console.log(`\n${pass} passed, ${fail} failed`);
  await browser.close();
})();
