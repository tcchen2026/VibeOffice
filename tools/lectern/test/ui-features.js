/* UI checks for the features added in this round: opening a password-protected deck (with a wrong password
 * first), keeping the password on save, Tools ▸ Options ▸ Security, editing an imported chart (dialog preview,
 * chart type change, saving from the model), and screenshots of imported charts.
 * Usage: node ui-features.js <outDir>   (static server as for corpus.js) */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const out = process.argv[2] || 'ui-out';
fs.mkdirSync(out, { recursive: true });
const HOST = process.env.HOST || 'http://127.0.0.1:8760/lectern/';
const C = process.env.CORPUS_URL || 'http://127.0.0.1:8766/pdata/corpus/';
const checks = [];
const ok = (name, cond, info) => { checks.push({ name, ok: !!cond, info }); console.log((cond ? 'PASS ' : 'FAIL ') + name + (info ? '  ' + JSON.stringify(info) : '')); };

(async () => {
  const b = await chromium.launch();
  const page = await b.newPage({ viewport: { width: 1360, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e.message)));
  await page.goto(HOST);
  await page.waitForFunction(() => window.L && L.app && L.pptx);
  await page.waitForTimeout(500);
  const open = (file) => page.evaluate(async ([url, name]) => {
    const buf = await (await fetch(url)).arrayBuffer();
    L.hist.dirty = false;
    window.__prev = L.pres;
    L.app.openFile(new File([buf], name));
  }, [C + encodeURIComponent(file), file.replace(/^[^_]+__[0-9a-f]+__/, '')]);

  /* 1. password-protected deck: wrong password, then the right one */
  await open('tika__5479add191f2__testPPT_protected_passtika.pptx');
  await page.waitForSelector('#pw-in', { timeout: 20000 });
  ok('password dialog appears', true);
  await page.fill('#pw-in', 'nope');
  await page.keyboard.press('Enter');
  await page.waitForFunction(() => document.querySelector('#pw-in') && /incorrect/.test(document.querySelector('.dlg-body').textContent), null, { timeout: 30000 });
  ok('wrong password is reported and asked again', true);
  await page.fill('#pw-in', 'tika');
  await page.keyboard.press('Enter');
  await page.waitForFunction(() => L.pres && L.pres.slides.length && /Encrypted/.test(L.model.allText(L.pres.slides[0])), null, { timeout: 30000 });
  const pw = await page.evaluate(() => ({ pw: L.pres.password, slides: L.pres.slides.length }));
  ok('encrypted deck opens with the right password', pw.slides === 1 && pw.pw === 'tika', pw);
  await page.screenshot({ path: path.join(out, '01-encrypted-open.png') });
  const saved = await page.evaluate(async () => {
    const u8 = new Uint8Array(await (await L.pptx.write(L.pres, { format: 'pptx' })).arrayBuffer());
    let needs = false; try { await L.pptx.read(u8.buffer.slice(0)); } catch (e) { needs = e.code === 'password'; }
    return { ole: u8[0] === 0xd0 && u8[1] === 0xcf, needs };
  });
  ok('saved again encrypted with the same password', saved.ole && saved.needs, saved);

  /* 2. Tools ▸ Options ▸ Security sets a new password */
  await open('oxsdk__e255bb9bc6d4__typical.pptx');
  await page.waitForFunction(() => L.pres !== window.__prev && L.pres.slides.length === 15, null, { timeout: 30000 });
  await page.evaluate(() => L.ui.exec('optionsDlg'));
  await page.waitForSelector('.dlg');
  await page.click('text=Security');
  await page.fill('#op-pw', 'Secret-2003');
  await page.screenshot({ path: path.join(out, '02-options-security.png') });
  await page.click('.dlg-foot >> text=OK');
  await page.waitForSelector('#pw-confirm');
  await page.fill('#pw-confirm', 'Secret-2003');
  await page.click('.dlg-foot >> text=OK');
  await page.waitForTimeout(300);
  const pw2 = await page.evaluate(() => L.pres.password);
  ok('Options ▸ Security sets the password to open', pw2 === 'Secret-2003', pw2);

  /* 3. imported chart: render, edit dialog, change type, save from the model */
  await open('automizer__8f1b9d60f455__SlideWithCharts.pptx');
  await page.waitForFunction(() => L.pres !== window.__prev && L.pres.slides.length > 1, null, { timeout: 30000 });
  await page.evaluate(() => L.ed.goto(1));
  await page.waitForTimeout(400);
  await page.screenshot({ path: path.join(out, '03-imported-charts.png') });
  const chartInfo = await page.evaluate(() => {
    const s = L.ed.slide();
    const ch = s.shapes.filter((x) => x.type === 'chart');
    return ch.map((c) => ({ id: c.id, kind: c.chart.kind, v: c.chart.v, srcId: !!c.chart.srcId, svgBars: document.querySelectorAll(`[data-sid="${c.id}"] svg rect`).length }));
  });
  ok('imported charts render from the v2 model', chartInfo.length === 2 && chartInfo.every((c) => c.v === 2 && c.srcId && c.svgBars > 0), chartInfo);
  const col = chartInfo.find((c) => c.kind === 'col');
  await page.evaluate((id) => { L.ed.select([id]); L.ui.exec('editChart'); }, col.id);
  await page.waitForSelector('.chart-dlg .chart-preview svg');
  await page.screenshot({ path: path.join(out, '04-chart-dialog.png') });
  await page.selectOption('.chart-dlg select >> nth=0', 'barStacked');
  await page.waitForTimeout(200);
  await page.screenshot({ path: path.join(out, '05-chart-dialog-stacked-bar.png') });
  await page.click('.dlg-foot >> text=OK');
  await page.waitForTimeout(400);
  const edited = await page.evaluate((id) => { const c = L.model.shapeById(L.ed.slide(), id).chart; return { kind: c.kind, edited: c.edited, groups: c.groups.map((g) => [g.type, g.dir, g.grouping]) }; }, col.id);
  ok('chart type changed through the dialog', edited.kind === 'barStacked' && edited.edited && edited.groups[0][1] === 'bar' && edited.groups[0][2] === 'stacked', edited);
  await page.screenshot({ path: path.join(out, '06-after-edit.png') });
  const written = await page.evaluate(async () => {
    const u8 = new Uint8Array(await (await L.pptx.write(L.pres, { format: 'pptx' })).arrayBuffer());
    const p2 = await L.pptx.read(u8.buffer);
    const kinds = []; for (const s of p2.slides) for (const sh of s.shapes) if (sh.type === 'chart') kinds.push(sh.chart.kind);
    let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
    return { kinds, b64: btoa(s) };
  });
  fs.writeFileSync(path.join(out, 'edited-chart.pptx'), Buffer.from(written.b64, 'base64'));
  ok('edited chart saved and read back as a stacked bar chart', written.kinds.includes('barStacked') && written.kinds.includes('pie'), written.kinds);

  /* 4. picture with a transparent colour and theme background texture */
  await open('locore__8eab9efe1181__tdf113163.pptx');
  await page.waitForFunction(() => L.pres !== window.__prev && L.pres.slides.length, null, { timeout: 30000 });
  await page.waitForTimeout(500);
  await page.screenshot({ path: path.join(out, '07-transparent-colour.png') });
  const clear = await page.evaluate(() => { let r = null; L.model.walk(L.pres.slides[0].shapes, (s) => { if (s.img && s.img.clear) r = { clear: s.img.clear, view: !!s.img.view }; return true; }); return r; });
  ok('picture keeps its transparent colour (a:clrChange)', clear && clear.clear === '#FFFFFF' && clear.view, clear);

  ok('no page errors', errors.length === 0, errors.slice(0, 5));
  fs.writeFileSync(path.join(out, 'checks.json'), JSON.stringify(checks, null, 1));
  console.log(`${checks.filter((c) => c.ok).length}/${checks.length} passed`);
  await b.close();
})().catch((e) => { console.error(e); process.exit(1); });
