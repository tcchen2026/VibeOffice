// Pictures of each page as Quire draws it in Print Layout at 100% (the page's own DOM, not the PDF export), for
// comparing with Word's PDFs of the same documents (python3 -I tools/lectern-vs-powerpoint.py wordDir outDir cmpDir: it
// compares any Office PDF with pictures named <name>__<page>.png):
//   node tools/quire-pictures.mjs docs.json outDir [nameFilter]      docs.json: [{ "name", "file" }]
// BASE as for the tests (another server, e.g. the code before a change).
import fs from 'node:fs';
import { ensureServer, openPage } from './ooxml/cdp.mjs';

const [list, out, only] = process.argv.slice(2);
if (!list || !out) { console.error('usage: node tools/quire-pictures.mjs docs.json outDir [nameFilter]'); process.exit(2); }
const docs = JSON.parse(fs.readFileSync(list, 'utf8')).filter((d) => !only || d.name.includes(only));
fs.mkdirSync(out, { recursive: true });
const stop = await ensureServer(process.env.BASE);
let page = null, used = 0;
try {
  for (const d of docs) {
    if (!page || used++ >= 25) {
      if (page) await page.close().catch(() => {});
      page = await openPage('quire', { persistence: false, automated: true, timeout: 120000 });
      await page.send('Emulation.setDeviceMetricsOverride', { width: 1400, height: 2200, deviceScaleFactor: 1, mobile: false });
      used = 1;
    }
    let n = 0;
    try {
      const b64 = fs.readFileSync(d.file).toString('base64');
      n = await page.evaluate(`(async () => {
        VO.opened = () => {}; L.layout.view = 'print'; L.layout.setZoom(1);
        const bytes = Uint8Array.from(atob(${JSON.stringify(b64)}), (c) => c.charCodeAt(0));
        const { doc } = await L.docx.read(bytes.buffer);
        L.app.loadDoc(doc, 'doc', { saved: true });
        await document.fonts.ready; L.layout.render();
        let last = -1, same = 0;
        for (let k = 0; k < 120 && same < 3; k++) { await new Promise((r) => setTimeout(r, 250)); const n = L.layout.pages.length; same = n === last ? same + 1 : 0; last = n; }
        /* as printed: no corner marks, nothing of the window over the pages, and no scroll bar to come and go and
           move the pages sideways while they are taken */
        if (!document.getElementById('__qp')) document.head.appendChild(Object.assign(document.createElement('style'), { id: '__qp', textContent: '.pg .crop{display:none!important} .cmdbars,.drawrow,.statusbar,.taskpane,.ruler-h,.ruler-v,.r-corner{visibility:hidden!important} .pg-hdr,.pg-ftr{opacity:1!important} *{scrollbar-width:none!important}' }));
        document.activeElement && document.activeElement.blur && document.activeElement.blur();
        getSelection().removeAllRanges();
        await document.fonts.ready;
        return L.layout.pages.length;
      })()`);
      /* a window wide and tall enough for the largest page, so none is cut off by the edge of the view */
      const big = await page.evaluate(`JSON.stringify(L.layout.pages.reduce((m, p) => [Math.max(m[0], p.el.offsetWidth), Math.max(m[1], p.el.offsetHeight)], [0, 0]))`);
      const [bw, bh] = JSON.parse(big);
      if (bw > 1200 || bh > 1900) {
        await page.send('Emulation.setDeviceMetricsOverride', { width: Math.max(1400, Math.ceil(bw) + 300), height: Math.max(2200, Math.ceil(bh) + 400), deviceScaleFactor: 1, mobile: false });
        await page.evaluate('new Promise((r) => setTimeout(r, 500))');
        used = 99;   // a fresh tab, at the usual size, for the next document
      }
      for (let i = 0; i < n; i++) {
        /* measured once the view has stopped moving */
        await page.evaluate(`L.layout.pages[${i}].el.scrollIntoView({ block: 'start', behavior: 'instant' })`);
        await new Promise((res) => setTimeout(res, 150));
        const where = `(() => { const q = L.layout.pages[${i}].el.getBoundingClientRect(); return { x: q.left + scrollX, y: q.top + scrollY, w: q.width, h: q.height }; })()`;
        /* taken again if the page moved meanwhile (something in the window can come or go and shift it sideways) */
        let r = await page.evaluate(where), shot;
        for (let k = 0; k < 3; k++) {
          shot = await page.send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: false, clip: { x: r.x, y: r.y, width: r.w, height: r.h, scale: 1 } });
          const again = await page.evaluate(where);
          if (Math.abs(again.x - r.x) < 0.5 && Math.abs(again.y - r.y) < 0.5) break;
          r = again;
        }
        fs.writeFileSync(`${out}/${d.name}__${i + 1}.png`, Buffer.from(shot.data, 'base64'));
      }
      console.log(d.name, n);
    } catch (e) {
      console.log(d.name, 'failed', String(e.message || e).slice(0, 120));
      await page.close().catch(() => {}); page = null;
    }
  }
} finally { if (page) await page.close().catch(() => {}); stop(); }
