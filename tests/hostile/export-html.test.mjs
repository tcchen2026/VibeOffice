// HTML we write from a document is a document of its own, opened from disk away from the suite's policy
// (todos/HARDENING.md 3.2): a hostile document's font names, colours, notes and link addresses go into it as text,
// never as markup, and the page carries its own policy (L.exportCSP). Each app's Save as Web Page, and the HTML
// Quire and Ledger put on the clipboard, is parsed for markup the writer didn't make; the pages are then opened
// in a tab, where nothing of the document's runs and Lectern's own script still does.
//
//   node --test --test-concurrency=1 tests/hostile/export-html.test.mjs
import { test, after } from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const pause = (ms) => new Promise((r) => setTimeout(r, ms));
const RAN = "document.title='ran'";
const FONT = `x"><img src=x onerror="${RAN}">`;              // a font name that ends the style attribute
const LINK = `https://example.com/" onmouseover="${RAN}`;      // a web address that ends the href
const SCRIPT_LINK = `javascript:${RAN}`;
const NOTE = `</script><script>${RAN}</script>`;              // a note that ends the page's script

/** what a parsed page holds that the writer never makes */
const inspect = (page, html) => page.evaluate(`(() => { const d = new DOMParser().parseFromString(${JSON.stringify(html)}, 'text/html');
  const all = [...d.querySelectorAll('*')];
  return { handlers: all.flatMap((e) => [...e.attributes].filter((a) => /^on/i.test(a.name)).map((a) => e.localName + '[' + a.name + ']')),
    pictures: d.querySelectorAll('img[src="x"]').length, scripts: d.querySelectorAll('script').length,
    scriptLinks: d.querySelectorAll('a[href^="javascript:" i]').length,
    policy: d.querySelector('meta[http-equiv="Content-Security-Policy"]')?.content.match(/script-src [^;]+/)?.[0] ?? null,
    fontAsText: all.some((e) => e.style && e.style.fontFamily.includes('onerror')),
    text: d.body.textContent.replace(/\\s+/g, ' ').trim() }; })()`);

/** the exported pages, served from 127.0.0.1 as files would be opened */
const pages = new Map();
const server = http.createServer((q, s) => { s.writeHead(pages.has(q.url) ? 200 : 404, { 'Content-Type': 'text/html; charset=utf-8' }); s.end(pages.get(q.url) || ''); });
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const stop = await ensureServer();
after(() => { server.close(); server.closeAllConnections(); stop(); });   // the browser's idle preconnections too
/** opens `html` in the test's tab: the title it ends with, and `probe` evaluated there */
async function visit(page, name, html, probe = 'null') {
  pages.set('/' + name, html);
  await page.send('Page.navigate', { url: `http://127.0.0.1:${server.address().port}/${name}` });
  await pause(800);
  return JSON.parse(await page.evaluate(`JSON.stringify({ title: document.title, probe: ${probe} })`));
}

test('Quire: Save as Web Page and copied HTML keep font names, colours and links as text', async () => {
  const page = await openPage('quire', { persistence: false, automated: true });
  try {
    await page.evaluate(`L.app.openFile(new File(['Some [Web](https://example.com/) and [Bad](https://example.com/bad) text\\n'], 'hostile.md'))`);
    await pause(800);
    await page.evaluate(`(() => { const runs = L.D.doc.main.blocks[0].runs; runs[0].rPr.font = ${JSON.stringify(FONT)};
      runs[3].rPr.link = { url: ${JSON.stringify(SCRIPT_LINK)} }; L.D.doc.bg = ${JSON.stringify(`fff}</style><script>${RAN}</script><style>`)}; })()`);
    const html = await page.evaluate(`L.htmlio.webPage(L.D.doc, 'hostile').then((b) => b.text())`);
    const clip = await page.evaluate(`L.htmlio.blocksToHTML(L.D.doc.main.blocks, { clipboard: true })`);
    const clean = { handlers: [], pictures: 0, scripts: 0, scriptLinks: 0, fontAsText: true, text: 'Some Web and Bad text' };
    assert.deepEqual(await inspect(page, html), { ...clean, policy: "script-src 'none'" }, 'the web page');
    assert.deepEqual(await inspect(page, clip), { ...clean, policy: null }, 'the clipboard');
    assert.equal((await visit(page, 'quire.htm', html)).title, 'hostile', 'opened, nothing of the document ran');
  } finally { await page.close(); }
});

test('Ledger: Save as Web Page and copied HTML keep font names and link addresses as text', async () => {
  const page = await openPage('ledger', { persistence: false, automated: true });
  try {
    const x = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;');
    const xml = `<?xml version="1.0"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
<Styles><Style ss:ID="s1"><Font ss:FontName="${x(FONT)}"/></Style></Styles>
<Worksheet ss:Name="Sheet1"><Table><Row><Cell ss:StyleID="s1" ss:HRef="${x(LINK)}"><Data ss:Type="String">Linked</Data></Cell></Row></Table></Worksheet></Workbook>`;
    await page.evaluate(`L.app.openFile(new File([${JSON.stringify(xml)}], 'hostile.xml'))`);
    await pause(800);
    const html = await page.evaluate(`L.csv.writeHTML(L.grid.sheet().wb).text()`);
    const clip = await page.evaluate(`(() => { const out = {}; L.clip.copy({ clipboardData: { setData: (t, v) => { out[t] = v; } } }); return out['text/html']; })()`);
    const clean = { handlers: [], pictures: 0, scripts: 0, scriptLinks: 0, fontAsText: true, text: 'Linked' };
    assert.deepEqual(await inspect(page, html), { ...clean, text: 'Sheet1Linked', policy: "script-src 'none'" }, 'the web page');
    assert.deepEqual(await inspect(page, clip), { ...clean, policy: null }, 'the clipboard');
    assert.equal(await page.evaluate(`new DOMParser().parseFromString(${JSON.stringify(html)}, 'text/html').querySelector('a').getAttribute('href')`), LINK,
      'the address is the link\'s, whole');
    assert.equal((await visit(page, 'ledger.htm', html)).title, 'hostile', 'opened, nothing of the document ran');
  } finally { await page.close(); }
});

test('Lectern: Save as Web Page keeps notes as data, and its own script still runs under its policy', async () => {
  const page = await openPage('lectern', { persistence: false, automated: true });
  try {
    await page.evaluate(`(() => { L.pres.slides[0].notes = ${JSON.stringify(NOTE)}; L.pres.props.title = 'hostile'; })()`);
    const html = await page.evaluate(`L.app.buildWebPage().then((b) => b.text())`);
    const found = await inspect(page, html);
    assert.deepEqual({ handlers: found.handlers, scripts: found.scripts, hashed: /^script-src 'sha256-[\w+/]+=*'$/.test(found.policy) },
      { handlers: [], scripts: 1, hashed: true }, 'one script, ours, allowed by its hash');
    assert.deepEqual(await visit(page, 'lectern.htm', html, `[typeof N === 'undefined' ? null : N[0], document.getElementById('num').textContent]`),
      { title: 'hostile', probe: [NOTE, '1 / 1'] }, 'the note is the script\'s data; the page shows its slide');
  } finally { await page.close(); }
});
