/* Real-deck test: in headless Chromium, open each .pptx with Lectern, load it into the application, render
 * every slide, rasterise the first slides to PNG, save it back to .pptx and open the saved copy again.
 * Usage: node corpus.js <listFile|dir> <out.jsonl> [--png outDir] [--save outDir] [--shard k/n]
 * Needs Lectern at HOST (default: tools/serve.py, http://127.0.0.1:8760/lectern/) and the decks at
 * CORPUS_URL (default http://127.0.0.1:8766/pdata/corpus/: python3 -m http.server 8766 in the folder above
 * the corpus). */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const args = process.argv.slice(2);
const opt = (k) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : null; };
const src = args[0], outFile = args[1];
const pngDir = opt('--png'), saveDir = opt('--save');
const shard = opt('--shard');
const HOST = process.env.HOST || 'http://127.0.0.1:8760/lectern/';
const CORPUS_URL = process.env.CORPUS_URL || 'http://127.0.0.1:8766/pdata/corpus/';
const PNG_SLIDES = +(process.env.PNG_SLIDES || 6), PNG_W = +(process.env.PNG_W || 480);
let files = fs.statSync(src).isDirectory() ? fs.readdirSync(src).filter((f) => /\.(pptx|pptm|ppsx|potx|ppsm|potm)$/i.test(f)).sort() : fs.readFileSync(src, 'utf8').split('\n').map((s) => s.trim()).filter(Boolean);
if (shard) { const [k, n] = shard.split('/').map(Number); files = files.filter((_, i) => i % n === k); }
const done = new Set();
if (fs.existsSync(outFile)) for (const l of fs.readFileSync(outFile, 'utf8').split('\n')) { try { done.add(JSON.parse(l).file); } catch (e) { /* partial line */ } }
files = files.filter((f) => !done.has(f));
for (const d of [pngDir, saveDir]) if (d) fs.mkdirSync(d, { recursive: true });

/* runs inside the page */
async function testDeck([url, o]) {
  const res = {};
  const warns = [];
  const ow = console.warn, oe = console.error;
  const fmt = (a) => a.map((x) => (x && x.stack ? x.stack.split('\n').slice(0, 3).join(' | ') : String(x))).join(' ').slice(0, 400);
  console.warn = (...a) => warns.push('W ' + fmt(a));
  console.error = (...a) => warns.push('E ' + fmt(a));
  const L = window.L, A = L.app, T = L.txt;
  const textOf = (sh) => {
    let t = '';
    if (sh.tx && sh.tx.ps) t += sh.tx.ps.map((p) => T.paraText(p)).join('\n');
    if (sh.tbl) for (const r of sh.tbl.rows) for (const c of r.cells) if (c && c.tx && c.tx.ps) t += '\n' + c.tx.ps.map((p) => T.paraText(p)).join('\n');
    if (sh.shapes) for (const k of sh.shapes) t += '\n' + textOf(k);
    return t;
  };
  const stats = (pres) => {
    const count = {}; let chars = 0;
    const per = [];
    const walk = (shs) => { for (const s of shs) { count[s.type] = (count[s.type] || 0) + 1; if (s.shapes) walk(s.shapes); } };
    for (const s of pres.slides) { walk(s.shapes); const t = s.shapes.map(textOf).join('\n'); chars += t.replace(/\s+/g, '').length; per.push({ n: s.shapes.length, t: t.replace(/\s+/g, ' ').trim() }); }
    return { count, chars, per };
  };
  try {
    const buf = await (await fetch(url)).arrayBuffer();
    res.size = buf.byteLength;
    let t0 = performance.now();
    let pres;
    try { pres = await L.pptx.read(buf); } catch (e) {
      /* password-protected decks: try the passwords their test suites document */
      if (e && e.code === 'password') for (const pw of o.passwords || []) { try { pres = await L.pptx.read(buf, { password: pw }); res.password = pw; break; } catch (e2) { if (e2.code !== 'badpassword') break; } }
      if (!pres) { res.readError = (e && e.code ? e.code + ': ' : '') + String((e && e.message) || e).slice(0, 300); return res; }
    }
    res.readMs = Math.round(performance.now() - t0);
    res.slides = pres.slides.length; res.W = pres.W; res.H = pres.H;
    if (pres.repaired) res.repaired = pres.repaired.parts ? pres.repaired.parts.length : true;
    const s1 = stats(pres);
    res.count = s1.count; res.chars = s1.chars;
    /* the application view: thumbnails, slide pane, notes */
    t0 = performance.now();
    try { A.loadPres(pres, 'deck', { saved: true }); } catch (e) { res.loadError = String((e && e.stack) || e).slice(0, 400); }
    res.loadMs = Math.round(performance.now() - t0);
    /* render every slide */
    t0 = performance.now();
    for (let i = 0; i < pres.slides.length; i++) {
      try { L.render.slide(pres, pres.slides[i], { mode: 'thumb', index: i }); } catch (e) { (res.renderErrors = res.renderErrors || []).push(i + ': ' + String((e && e.stack) || e).split('\n').slice(0, 3).join(' | ').slice(0, 300)); }
    }
    res.renderMs = Math.round(performance.now() - t0);
    if (o.png) {
      res.png = [];
      for (let i = 0; i < Math.min(o.pngSlides, pres.slides.length); i++) {
        try { const c = await A.slideCanvas(pres.slides[i], o.pngW); res.png.push(c.toDataURL('image/png')); } catch (e) { res.png.push(null); (res.pngErrors = res.pngErrors || []).push(i + ': ' + String((e && e.message) || e).slice(0, 200)); }
      }
    }
    /* save and reopen */
    t0 = performance.now();
    let out;
    try { const blob = await L.pptx.write(pres, { format: 'pptx' }); out = new Uint8Array(await blob.arrayBuffer()); } catch (e) { res.writeError = String((e && e.stack) || e).split('\n').slice(0, 4).join(' | ').slice(0, 400); return res; }
    res.writeMs = Math.round(performance.now() - t0);
    res.outSize = out.length;
    let p2;
    try { p2 = await L.pptx.read(out.buffer, { password: pres.password }); } catch (e) { res.rereadError = String((e && e.message) || e).slice(0, 300); return res; }
    if (p2.repaired) res.rereadRepaired = true;
    const s2 = stats(p2);
    const diffs = [];
    if (p2.slides.length !== pres.slides.length) diffs.push('slides ' + pres.slides.length + ' vs ' + p2.slides.length);
    for (let i = 0; i < Math.min(s1.per.length, s2.per.length); i++) {
      if (s1.per[i].n !== s2.per[i].n) diffs.push(`slide ${i + 1} shapes ${s1.per[i].n} vs ${s2.per[i].n}`);
      if (s1.per[i].t !== s2.per[i].t) diffs.push(`slide ${i + 1} text ${JSON.stringify(s1.per[i].t.slice(0, 80))} vs ${JSON.stringify(s2.per[i].t.slice(0, 80))}`);
    }
    for (const k of new Set(Object.keys(s1.count).concat(Object.keys(s2.count)))) if ((s1.count[k] || 0) !== (s2.count[k] || 0)) diffs.push(`${k} ${s1.count[k] || 0} vs ${s2.count[k] || 0}`);
    if (diffs.length) res.rtDiffs = diffs.slice(0, 12), res.rtDiffCount = diffs.length;
    if (o.save) { let s = ''; for (let i = 0; i < out.length; i += 0x8000) s += String.fromCharCode.apply(null, out.subarray(i, i + 0x8000)); res.saved = btoa(s); }
  } catch (e) {
    res.fatal = String((e && e.stack) || e).split('\n').slice(0, 4).join(' | ').slice(0, 400);
  } finally {
    console.warn = ow; console.error = oe;
    if (warns.length) { res.warnCount = warns.length; res.warns = Array.from(new Set(warns)).slice(0, 12); }
  }
  return res;
}

(async () => {
  const browser = await chromium.launch({ args: ['--disable-web-security'] });
  let ctx, page, pageErrors = [];
  const fresh = async () => {
    if (ctx) await ctx.close().catch(() => {});
    ctx = await browser.newContext({ viewport: { width: 1280, height: 860 } });
    page = await ctx.newPage();
    pageErrors = [];
    page.on('pageerror', (e) => pageErrors.push(String(e.message).slice(0, 300)));
    page.on('dialog', (d) => d.dismiss().catch(() => {}));
    await page.goto(HOST, { waitUntil: 'load' });
    await page.waitForFunction(() => window.L && L.app && L.pptx && L.pptx.read, null, { timeout: 30000 });
  };
  await fresh();
  const out = fs.createWriteStream(outFile, { flags: 'a' });
  let n = 0;
  for (const f of files) {
    n++;
    const t0 = Date.now();
    let res;
    pageErrors = [];
    try {
      res = await Promise.race([
        page.evaluate(testDeck, [CORPUS_URL + encodeURIComponent(f), { png: !!pngDir, pngSlides: PNG_SLIDES, pngW: PNG_W, save: !!saveDir, passwords: (process.env.PASSWORDS || 'pass,tika').split(',') }]).then((r) => r),
        new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), +(process.env.TIMEOUT || 180000))),
      ]);
    } catch (e) {
      res = { fatal: String(e.message || e).slice(0, 300) };
      await fresh().catch(() => {});
    }
    res = Object.assign({ file: f }, res, { ms: Date.now() - t0 });
    if (pageErrors.length) res.pageErrors = pageErrors.slice(0, 6);
    if (res.png) { res.png.forEach((d, i) => { if (d) fs.writeFileSync(path.join(pngDir, f.replace(/\.[^.]+$/, '') + '__' + (i + 1) + '.png'), Buffer.from(d.split(',')[1], 'base64')); }); res.pngs = res.png.filter(Boolean).length; delete res.png; }
    if (res.saved) { fs.writeFileSync(path.join(saveDir, f.replace(/\.[^.]+$/, '') + '.pptx'), Buffer.from(res.saved, 'base64')); delete res.saved; }
    out.write(JSON.stringify(res) + '\n');
    if (n % 25 === 0) { process.stderr.write(`${n}/${files.length}\n`); await fresh(); }
  }
  out.end();
  await browser.close();
})();
