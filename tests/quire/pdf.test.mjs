// Quire's PDF text comes from the painted DOM, including text split across styled runs.
// Chromium measures the ranges; Poppler independently checks the exported document's words.
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

let page, stop, dir;
before(async () => {
  stop = await ensureServer();
  page = await openPage('quire', { persistence: false, automated: true });
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'quire-pdf-'));
});
after(async () => {
  if (page) await page.close();
  stop?.();
  if (dir) fs.rmSync(dir, { recursive: true, force: true });
});

async function collect(html, zoom = 1) {
  return page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true);
    if (!pdf.collectDOM) throw new Error('The PDF DOM collector has not loaded');
    const el = document.createElement('div');
    el.style.cssText = 'position:absolute;left:0;top:0;width:480px;height:700px;' +
      'font:16px/24px Arial;color:black;background:white;transform-origin:0 0;' +
      'transform:scale(${zoom});';
    el.innerHTML = ${JSON.stringify(html)};
    document.body.append(el);
    try {
      await document.fonts.ready;
      const before = el.outerHTML;
      const lines = pdf.collectDOM(el, ${0.75 / zoom});
      return { lines, unchanged: before === el.outerHTML };
    } finally { el.remove(); }
  })()`);
}
const rows = result => result.lines.map(line => line.map(word => word.text).join(' '));

test('PDF collector joins styled word pieces and punctuation, with distinct lines and zoom-independent boxes', async () => {
  const html = '<div>Hel<b>lo</b>, <i>world</i>!</div><div>Next line</div>';
  const full = await collect(html), zoomed = await collect(html, 0.6);
  assert.deepEqual(rows(full), ['Hello, world!', 'Next line']);
  assert.deepEqual(rows(zoomed), rows(full));
  for (let i = 0; i < full.lines.length; i++) for (let j = 0; j < full.lines[i].length; j++) {
    for (const key of ['x', 'y', 'w', 'h']) {
      assert.ok(Math.abs(full.lines[i][j][key] - zoomed.lines[i][j][key]) < 0.1, key + ' at zoom 0.6');
    }
  }
  assert.ok(full.unchanged && zoomed.unchanged);
});

test('PDF collector keeps wrapped CJK and Unicode scalars, and separates columns and table cells in DOM order', async () => {
  const cjk = '日本語漢字中文日本語漢字中文日本語漢字中文';
  const result = await collect('<div style="width:55px;overflow-wrap:anywhere">' + cjk + '</div>' +
    '<div>Emoji 😀 rare 𠮷</div>' +
    '<div style="position:relative;height:55px"><div style="position:absolute;left:0;top:0">' +
    '<div>Left first</div><div>Left second</div></div><div style="position:absolute;left:180px;top:0">' +
    '<div>Right first</div><div>Right second</div></div></div>' +
    '<table><tbody><tr><td>Cell one</td><td>Cell two</td></tr></tbody></table>');
  const text = rows(result), end = text.findIndex(line => line.startsWith('Emoji'));
  assert.ok(end > 1, 'the CJK string wraps');
  assert.equal(text.slice(0, end).join(''), cjk);
  assert.deepEqual(text.slice(end), ['Emoji 😀 rare 𠮷', 'Left first', 'Left second',
    'Right first', 'Right second', 'Cell one', 'Cell two']);
});

test('PDF collector omits hidden decorations and clipped letters, but includes off-screen page text', async () => {
  const result = await collect('<div>Visible text</div>' +
    '<div class="format-marks">FORMATTING</div><div class="sectmark">SECTION</div>' +
    '<div class="hbrk">BREAK</div><div class="crop">CROP</div><div class="balloons">BALLOON</div>' +
    '<div style="display:none">DISPLAY</div><div style="opacity:0">TRANSPARENT</div>' +
    '<div aria-hidden="true">DECORATION</div>' +
    '<div style="visibility:hidden">HIDDEN<span style="visibility:visible">Restored</span></div>' +
    '<div style="width:62px;height:24px;overflow:hidden;font:16px/24px monospace;white-space:nowrap">' +
    '<span>ABCDEFGHIJKLMNO</span><div>CLIPPEDROW</div></div>' +
    '<div style="position:absolute;top:710px">OUTSIDE</div>');
  const text = rows(result).join('\n');
  assert.match(text, /Visible text\nRestored\nABCDEF/);
  assert.doesNotMatch(text, /FORMATTING|SECTION|BREAK|CROP|BALLOON|DISPLAY|TRANSPARENT|DECORATION|HIDDEN|IJKLMNO|CLIPPEDROW|OUTSIDE/);
  const offscreen = await collect('<div style="position:relative;top:500px">Still on page</div>');
  assert.deepEqual(rows(offscreen), ['Still on page']);
  assert.ok(result.lines.flat().every(word => word.x + word.w <= 360.01 && word.y + word.h <= 525.01));
});

test('PDF collector takes warped text once from its transparent source body, never the SVG slices', async () => {
  const result = await collect('<div class="shp warped" style="position:relative;width:200px;height:60px">' +
    '<svg class="warp" aria-hidden="true" width="200" height="60">' +
    '<text x="0" y="20">SLICES SLICES SLICES</text></svg>' +
    '<div class="tbx" style="left:0;top:0;width:200px;height:50px;opacity:0">Bent words</div></div>' +
    '<div class="shp warped" style="position:relative;width:200px;height:60px">' +
    '<div class="tx" style="opacity:0">Slide words</div></div>' +
    '<div style="opacity:0"><div class="shp warped"><div class="tbx">HIDDENSHAPE</div></div></div>');
  assert.deepEqual(rows(result), ['Bent words', 'Slide words']);
});

test('PDF collector measures SVG chart text and MathML while excluding definitions and annotations', async () => {
  const result = await collect('<svg xmlns="http://www.w3.org/2000/svg" width="300" height="60">' +
    '<defs><text id="unused">DEFINITION</text></defs><title>DESCRIPTION</title>' +
    '<text x="10" y="30">Chart label</text></svg>' +
    '<math xmlns="http://www.w3.org/1998/Math/MathML"><semantics><mrow><mi>x</mi><mo>+</mo>' +
    '<mn>2</mn></mrow><annotation encoding="text/plain">ANNOTATION</annotation></semantics></math>');
  const text = rows(result).join('\n');
  assert.match(text, /Chart label/);
  assert.equal(text.replace(/Chart label|\s/g, ''), 'x+2');
  assert.deepEqual(result.lines.flat().filter(word => !['Chart', 'label'].includes(word.text))
    .map(word => word.text), ['x', '+', '2'], 'each math token keeps its own selection box');
  assert.ok(result.lines.flat().every(word => word.w > 0 && word.h > 0));
});

test('PDF collector preserves the spaces represented by empty tabs and list-label spacers', async () => {
  const result = await collect('<div>Chapter One<span class="tb" style="width:80px"></span>7</div>' +
    '<div><span class="lbl">•</span><span class="lsp" style="display:inline-block;width:12px"></span>' +
    'First item</div>');
  assert.deepEqual(rows(result), ['Chapter One 7', '• First item']);
});

test('PDF copy keeps RTL word runs on each side of punctuation in their source order', async () => {
  const result = await collect('<div dir="rtl">الداونلود(يستعمل</div><div dir="rtl">الداونلود–في</div>');
  const data = await page.evaluate(`(async () => {
    const canvas = document.createElement('canvas');
    canvas.width = canvas.height = 1;
    const blob = await L.buildPDF([{ jpeg: L.pdf.jpeg(canvas), iw: 1, ih: 1, pw: 360, ph: 525,
      lines: ${JSON.stringify(result.lines)} }]);
    return (await L.blobToDataURL(blob)).split(',')[1];
  })()`);
  const file = path.join(dir, 'rtl.pdf');
  fs.writeFileSync(file, Buffer.from(data, 'base64'));
  const poppler = execFileSync('pdftotext', ['-raw', file, '-'], { encoding: 'utf8' });
  const pdfium = execFileSync('python3', ['-I', '-c',
    'import sys,pypdfium2 as pdf; d=pdf.PdfDocument(sys.argv[1]); print(d[0].get_textpage().get_text_range())', file],
  { encoding: 'utf8' });
  for (const text of [poppler, pdfium]) {
    const plain = text.replace(/[\u202a-\u202e\u2066-\u2069]/g, '');
    assert.match(plain, /الداونلود.يستعمل/);
    assert.match(plain, /الداونلود–في/);
  }
});

test('Quire PDF omits the hidden lines of an exact-height table cell', async () => {
  const result = await page.evaluate(`(async () => {
    const { D } = L, doc = D.newDoc(), table = D.simpleTable(1, 1, 220);
    table.rows[0].trPr = { h: 24, hRule: 'exact' };
    table.rows[0].cells[0].blocks = [D.para([
      D.text('Visible cell line'), D.item('br'), D.item('br'), D.text('HIDDENCELLLINE'),
    ], { spacing: { before: 0, after: 0, line: 240, rule: 'auto' } })];
    doc.main.blocks = [table];
    L.app.loadDoc(doc, 'Clipped cell', { saved: true });
    await document.fonts.ready;
    L.layout.render();
    const blob = await L.app.buildPDF();
    L.ui.busy(false);
    return (await L.blobToDataURL(blob)).split(',')[1];
  })()`);
  const file = path.join(dir, 'clipped-cell.pdf');
  fs.writeFileSync(file, Buffer.from(result, 'base64'));
  const text = execFileSync('pdftotext', ['-raw', file, '-'], { encoding: 'utf8' });
  assert.match(text, /Visible cell line/);
  assert.doesNotMatch(text, /HIDDENCELLLINE/);
});

test('Quire sample PDF extracts body text, tables, headers, footers and notes with spaces and line breaks', async () => {
  const sample = await openPage('quire', { persistence: false, automated: true });
  let result;
  try {
    result = await sample.evaluate(`(async () => {
      const loadedBefore = !!L.pdf?.collectDOM;
      L.app.loadDoc(L.app.buildSample(), 'PDF sample', { saved: true });
      await document.fonts.ready;
      L.layout.render();
      const blob = await L.app.buildPDF();
      L.ui.busy(false);
      const data = await L.blobToDataURL(blob);
      return { data: data.slice(data.indexOf(',') + 1), pages: L.layout.pages.length,
        loadedBefore, loadedAfter: !!L.pdf?.collectDOM };
    })()`);
    assert.deepEqual(sample.errors, []);
    assert.deepEqual(sample.requests, []);
  } finally { await sample.close(); }
  assert.equal(result.loadedBefore, false, 'the collector is not loaded on app startup');
  assert.equal(result.loadedAfter, true, 'the first export loads the collector');
  const file = path.join(dir, 'sample.pdf');
  fs.writeFileSync(file, Buffer.from(result.data, 'base64'));
  const text = execFileSync('pdftotext', ['-raw', file, '-'], { encoding: 'utf8' });
  assert.match(text, /Welcome to Quire\n/);
  assert.match(text.replace(/\s+/g, ' '), /bold, italic, underline, highlighting, colors, styles/);
  for (const content of ['Sample document', 'Find and Replace', 'Ctrl+Alt+F',
    'Footnotes are numbered automatically', 'Chapter One', 'Page 1 of']) {
    assert.ok(text.includes(content), content);
  }
  assert.equal((text.match(/\f/g) || []).length, result.pages);
  assert.deepEqual(page.errors, []);
  assert.deepEqual(page.requests, []);
});
