// Pictures of each slide as Lectern draws it on screen (the slide's DOM, not the image export), for comparing with
// PowerPoint's PDFs of the same decks (tools/lectern-vs-powerpoint.py):
//   node tools/lectern-pictures.mjs decks.json outDir [nameFilter]      decks.json: [{ "name", "file" }]
// writes outDir/<name>__<slide>.png, W px wide (env W, default 960: a point a pixel on a 13.33-inch slide); BASE as for
// the tests (another server, e.g. the code before a change).
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from './ooxml/cdp.mjs';

const [list, out, only] = process.argv.slice(2);
if (!list || !out) { console.error('usage: node tools/lectern-pictures.mjs decks.json outDir [nameFilter]'); process.exit(2); }
const decks = JSON.parse(fs.readFileSync(list, 'utf8')).filter((d) => !only || d.name.includes(only));
fs.mkdirSync(out, { recursive: true });
const W = +(process.env.W || 960);
const stop = await ensureServer(process.env.BASE);
const page = await openPage('lectern', { persistence: false, automated: true, timeout: 300000 });
await page.send('Emulation.setDeviceMetricsOverride', { width: Math.max(1400, W + 40), height: Math.max(1100, Math.round(W * 0.8)), deviceScaleFactor: 1, mobile: false });
try {
  for (const d of decks) {
    const b64 = fs.readFileSync(d.file).toString('base64');
    let n;
    try {
      n = await page.evaluate(`(async () => {
        const pres = await L.pptx.read(Uint8Array.from(atob(${JSON.stringify(b64)}), (c) => c.charCodeAt(0)));
        L.app.loadPres(pres, 'deck', { saved: true });
        let host = document.getElementById('__shot');
        if (!host) { host = document.createElement('div'); host.id = '__shot'; document.body.appendChild(host); }
        host.style.cssText = 'position:fixed;left:0;top:0;z-index:2147483647;background:#fff;transform-origin:0 0;';
        window.__pres = pres;
        return pres.slides.length;
      })()`);
    } catch (e) { console.log(d.name, 'not opened:', String(e.message).slice(0, 200)); continue; }
    for (let i = 0; i < n; i++) {
      const h = await page.evaluate(`(async () => {
        const pres = window.__pres, host = document.getElementById('__shot');
        host.textContent = '';
        host.appendChild(L.render.slide(pres, pres.slides[${i}], { mode: 'thumb', index: ${i} }));
        const k = ${W} / pres.W;
        host.style.width = pres.W + 'px'; host.style.height = pres.H + 'px'; host.style.transform = 'scale(' + k + ')';
        await Promise.all([...host.querySelectorAll('img')].map((im) => im.decode().catch(() => {})));
        await document.fonts.ready;
        await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
        return Math.round(pres.H * k);
      })()`);
      const shot = await page.send('Page.captureScreenshot', { format: 'png', clip: { x: 0, y: 0, width: W, height: h, scale: 1 } });
      fs.writeFileSync(path.join(out, d.name + '__' + (i + 1) + '.png'), Buffer.from(shot.data, 'base64'));
    }
    console.log(d.name, n, page.errors.length ? 'errors: ' + page.errors.splice(0).join(' | ').slice(0, 300) : '');
  }
} finally {
  await page.evaluate(`document.getElementById('__shot')?.remove()`).catch(() => {});
  await page.close(); stop();
}
