/* Execute every command once (closing dialogs / menus), report page errors. */
const { chromium } = require('playwright');
(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push('pageerror: ' + e.message + ' | ' + (e.stack || '').split('\n').slice(1, 4).join(' / ')));
  page.on('console', (m) => { if (m.type() === 'error' && !/ERR_TUNNEL|Failed to load resource/.test(m.text())) errors.push('console: ' + m.text()); });
  await page.goto('http://127.0.0.1:8760/ledger/');
  await page.waitForTimeout(800);
  const ids = await page.evaluate(() => Object.keys(L.ui.cmds));
  const skip = new Set(['open', 'insertPicture', 'insertClipArt', 'importText', 'sheetBackground', 'exitApp', 'close', 'newWindow', 'printQuick', 'exportPDF', 'exportCSV', 'saveWeb', 'save', 'print', 'fullScreen']);
  const results = [];
  for (const id of ids) {
    if (skip.has(id)) continue;
    if (process.env.V) console.error(id);
    const before = errors.length;
    try {
      await page.evaluate((id) => { window.__lastCmd = id; try { L.ui.exec(id); } catch (e) { console.error('exec ' + id + ': ' + e.message); } }, id);
      await page.waitForTimeout(60);
      /* close whatever opened */
      for (let k = 0; k < 4; k++) {
        const open = await page.evaluate(() => (L.ui.dialogOpen() || L.ui.menuOpen() || !!document.querySelector('.pp-overlay,.af-list,.cmt-edit,.dlg-overlay.modeless,.draw-overlay')));
        if (!open) break;
        await page.evaluate(() => { const m = document.querySelector('.dlg-overlay.modeless'); if (m) m.remove(); const p = document.querySelector('.pp-overlay'); if (p) { document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })); } });
        await page.keyboard.press('Escape');
        await page.waitForTimeout(40);
      }
      await page.evaluate(() => { document.querySelectorAll('.dlg-overlay, .cmt-edit, .draw-overlay, .af-list').forEach((x) => x.remove()); L.ui.closeMenus(); if (L.editor.active) L.editor.cancel(); });
    } catch (e) { errors.push('playwright ' + id + ': ' + e.message); }
    if (errors.length > before) results.push(id + ' → ' + errors.slice(before).join(' || '));
  }
  console.log(results.join('\n') || 'all commands ran without errors');
  console.log('commands:', ids.length);
  await browser.close();
})();
