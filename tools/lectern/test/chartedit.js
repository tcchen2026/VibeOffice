/* Charts edited in Lectern are written from the chart model: open each deck, mark every chart as edited
 * (optionally switching its type), save, and write the result for external validation.
 * Usage: node chartedit.js <listFile> <outDir> [kind] */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const [list, outDir, kind] = process.argv.slice(2);
const HOST = process.env.HOST || 'http://127.0.0.1:8760/lectern/';
const CORPUS_URL = process.env.CORPUS_URL || 'http://127.0.0.1:8766/pdata/corpus/';
fs.mkdirSync(outDir, { recursive: true });
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto(HOST);
  await p.waitForFunction(() => window.L && L.pptx && L.pptx.read);
  const files = fs.readFileSync(list, 'utf8').split('\n').filter(Boolean);
  let n = 0, charts = 0, errs = 0;
  for (const f of files) {
    const r = await p.evaluate(async ([url, kind]) => {
      try {
        const pres = await L.pptx.read(await (await fetch(url)).arrayBuffer());
        let n = 0;
        const walk = (shs) => { for (const s of shs) { if (s.type === 'chart' && s.chart) { const prev = s.chart.kind; if (kind) s.chart.kind = kind; L.chart.syncV2(s.chart, prev); s.chart.edited = true; L.chart.render(s.chart, s.w, s.h, null); n++; } if (s.kids) walk(s.kids); } };
        pres.slides.forEach((s) => walk(s.shapes));
        const blob = await L.pptx.write(pres, { format: 'pptx' });
        const u8 = new Uint8Array(await blob.arrayBuffer());
        const p2 = await L.pptx.read(u8.buffer);
        let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
        return { n, data: btoa(s), slides: p2.slides.length };
      } catch (e) { return { err: String(e && e.stack || e).slice(0, 300) }; }
    }, [CORPUS_URL + encodeURIComponent(f), kind || null]);
    if (r.err) { errs++; console.log('ERR', f, r.err); continue; }
    charts += r.n; n++;
    if (r.n) fs.writeFileSync(path.join(outDir, f.replace(/\.[^.]+$/, '') + '.pptx'), Buffer.from(r.data, 'base64'));
  }
  console.log('decks', n, 'charts', charts, 'errors', errs);
  await b.close();
})();
