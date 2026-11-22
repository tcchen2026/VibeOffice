// Every DrawingML preset shape (common/presets.js) at its default adjustments, filled and outlined, each in a labelled
// cell, twice (wide and tall cells), saved as tmp/shapes/shapes.pptx: open it in PowerPoint and save it as PDF, then compare
// PowerPoint's drawing with ours (tools/shapes/compare.py):  node tools/shapes/deck.mjs
import fs from 'node:fs';
import { ensureServer, openPage } from '../ooxml/cdp.mjs';

const stop = await ensureServer();
const page = await openPage('lectern', { automated: true, persistence: false });
try {
  const b64 = await page.evaluate(`(async () => {
    const M = L.model;
    const names = Object.keys(L.presetShapes).filter((k) => !/^(line|lineInv|straightConnector1|bentConnector\\d|curvedConnector\\d)$/.test(k)).sort();
    const pres = M.newPresentation({ empty: true, w: 960, h: 540 });
    const cols = 5, rows = 3, per = cols * rows;   // far enough apart that callout lines stay in their cells
    for (const [tall, w, h] of [[false, 110, 70], [true, 70, 100]]) {
      for (let s = 0; s * per < names.length; s++) {
        const slide = M.newSlide(pres, 'blank'); pres.slides.push(slide);
        names.slice(s * per, s * per + per).forEach((name, i) => {
          const cx = 20 + (i % cols) * 188, cy = 14 + Math.floor(i / cols) * 176;
          const x = cx + (184 - w) / 2, y = cy + 6 + (150 - h) / 2;
          const sh = M.newShape(pres, slide, 'rect', x, y, w, h);
          sh.geom = name; delete sh.adj;
          sh.fill = { t: 'solid', c: '#9DC3E6', a: 1 }; sh.line = { c: '#000000', w: 1, dash: 'solid' }; sh.name = name;
          if (sh.tx) sh.tx.ps = [L.txt.para('')];
          slide.shapes.push(sh);
          const tb = M.newTextBox(pres, slide, cx, cy + 158, 184, 16);
          tb.tx.ps = [L.txt.para(name, { algn: 'ctr' }, { sz: 8 })];
          tb.name = 'label ' + name; slide.shapes.push(tb);
        });
      }
    }
    const blob = await L.pptx.write(pres);
    const a = new Uint8Array(await blob.arrayBuffer()); let s = ''; for (let i = 0; i < a.length; i += 32768) s += String.fromCharCode(...a.subarray(i, i + 32768)); return btoa(s);
  })()`);
  fs.mkdirSync('tmp/shapes', { recursive: true });
  fs.writeFileSync('tmp/shapes/shapes.pptx', Buffer.from(b64, 'base64'));
  console.log('saved tmp/shapes/shapes.pptx', fs.statSync('tmp/shapes/shapes.pptx').size, 'bytes');
  if (page.errors.length) console.log('errors', page.errors);
} finally { await page.close(); stop(); }
