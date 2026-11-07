// Ledger's PDF text follows the actual canvas calls, without changing their pixels.
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
  page = await openPage('ledger', { persistence: false, automated: true });
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ledger-pdf-'));
});
after(async () => {
  if (page) await page.close();
  stop?.();
  if (dir) fs.rmSync(dir, { recursive: true, force: true });
});

test('canvas recording forwards paint, properties and methods without changing pixels', async () => {
  const result = await page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true);
    function draw(record) {
      const cv = document.createElement('canvas'); cv.width = 420; cv.height = 180;
      const raw = cv.getContext('2d'), ctx = record ? pdf.recordCanvas(raw) : raw;
      const gradient = ctx.createLinearGradient(0, 0, 420, 180);
      gradient.addColorStop(0, '#fff'); gradient.addColorStop(1, '#ddd');
      ctx.fillStyle = gradient; ctx.fillRect(0, 0, 420, 180);
      ctx.save(); ctx.translate(20, 30); ctx.rotate(.12); ctx.scale(1.2, 1.2);
      ctx.font = 'bold 24px Arial'; ctx.fillStyle = '#123456'; ctx.fillText('Hello canvas', 0, 30);
      ctx.restore(); ctx.font = '18px Arial'; ctx.fillStyle = '#000'; ctx.fillText('Second row', 20, 120);
      return { data: cv.toDataURL(), lines: record ? pdf.canvasLines(cv) : [],
        font: raw.font, canvas: ctx.canvas === cv };
    }
    const plain = draw(false), recorded = draw(true);
    return { same: plain.data === recorded.data, ...recorded, data: undefined };
  })()`);
  assert.equal(result.same, true);
  assert.equal(result.canvas, true);
  assert.equal(result.font, '18px Arial');
  assert.deepEqual(result.lines.map(line => line.map(w => w.text).join(' ')), ['Hello canvas', 'Second row']);
  assert.ok(Math.abs(result.lines[0][0].angle - .12 * 180 / Math.PI) < .001);
});

test('canvas word boxes follow alignment, font baselines, scaling and maxWidth', async () => {
  const rows = await page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true), rows = [];
    for (const baseline of ['alphabetic', 'top', 'middle', 'bottom', 'hanging', 'ideographic']) {
      for (const align of ['left', 'center', 'right', 'start', 'end']) {
        const cv = document.createElement('canvas'); cv.width = 900; cv.height = 450;
        const raw = cv.getContext('2d'), ctx = pdf.recordCanvas(raw);
        ctx.font = '20px Arial'; ctx.textAlign = align; ctx.textBaseline = baseline;
        ctx.direction = 'rtl'; ctx.translate(40, 30); ctx.scale(1.5, 1.2);
        const m = ctx.measureText('Wide word'), first = ctx.measureText('Wide').width;
        const limit = m.width * .6, right = align === 'right' || align === 'start';
        const left = 300 - (right ? limit : align === 'center' ? limit / 2 : 0);
        ctx.fillText('Wide word', 300, 100, limit);
        rows.push({ baseline, align, words: pdf.canvasLines(cv).flat(), expected: {
          x: 40 + left * 1.5, y: 30 + (100 - m.fontBoundingBoxAscent) * 1.2,
          w: first * .6 * 1.5, h: (m.fontBoundingBoxAscent + m.fontBoundingBoxDescent) * 1.2 } });
      }
    }
    return rows;
  })()`);
  for (const row of rows) {
    assert.deepEqual(row.words.map(w => w.text), ['Wide', 'word']);
    for (const key of ['x', 'y', 'w', 'h']) {
      assert.ok(Math.abs(row.words[0][key] - row.expected[key]) < .1,
        row.baseline + ' ' + row.align + ' ' + key + ': ' + row.words[0][key]);
    }
  }
});

test('canvas clips keep only visible graphemes through save/restore and rotated rectangles', async () => {
  const result = await page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true), cv = document.createElement('canvas');
    cv.width = 500; cv.height = 350;
    const ctx = pdf.recordCanvas(cv.getContext('2d'));
    ctx.font = '20px monospace'; ctx.textBaseline = 'top';
    const cw = ctx.measureText('A').width;
    ctx.save(); ctx.beginPath(); ctx.rect(20, 20, cw * 5.3, 25); ctx.clip();
    ctx.fillText('ABCDEFGHIJKLM', 20, 20); ctx.fillText('HIDDEN ROW', 20, 60);
    ctx.restore(); ctx.fillText('Outside clip', 20, 110);
    ctx.save(); ctx.translate(260, 100); ctx.rotate(Math.PI / 4);
    ctx.beginPath(); ctx.rect(0, 0, cw * 5.3, 25);
    ctx.save(); ctx.beginPath(); ctx.rect(0, 0, cw * 3.3, 25); ctx.restore();
    ctx.clip(); // The current path is not part of the saved drawing state.
    ctx.fillText('MNOPQRSTUV', 0, 0); ctx.restore();
    ctx.globalAlpha = 0; ctx.fillText('TRANSPARENT', 20, 160);
    return pdf.canvasLines(cv);
  })()`);
  assert.deepEqual(result.map(line => line.map(w => w.text).join(' ')), ['ABCDEF', 'Outside clip', 'MNOP']);
  assert.ok(Math.abs(result[2][0].angle - 45) < .001);
});

test('recorded child canvases retain text through drawImage crop, resize and parent clipping', async () => {
  const result = await page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true), child = document.createElement('canvas');
    child.width = 300; child.height = 90;
    const a = pdf.recordCanvas(child.getContext('2d'));
    a.font = '20px monospace'; a.textBaseline = 'top'; a.fillText('ABCDEFGHIJ', 10, 20);
    const cw = a.measureText('A').width, parent = document.createElement('canvas');
    parent.width = 500; parent.height = 300;
    const b = pdf.recordCanvas(parent.getContext('2d'));
    b.save(); b.beginPath(); b.rect(100, 60, 3.4 * cw * 2, 120); b.clip();
    b.drawImage(child, 10 + 2 * cw, 0, 5 * cw, 60, 100, 60, 10 * cw, 120);
    b.restore();
    const lines = pdf.canvasLines(parent, .5);
    return { lines, cw };
  })()`);
  assert.equal(result.lines.flat().map(w => w.text).join(' '), 'CDEF');
  const word = result.lines[0][0];
  assert.ok(Math.abs(word.x - 50) < .01);
  assert.ok(Math.abs(word.w - result.cw * 3.4) < .01);
});

test('canvas bidi ranges retain logical text, inherited direction and clipped Unicode graphemes', async () => {
  const result = await page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true), cv = document.createElement('canvas');
    cv.width = 600; cv.height = 180; cv.style.direction = 'rtl'; document.body.append(cv);
    try {
      const ctx = pdf.recordCanvas(cv.getContext('2d'));
      ctx.font = '24px Arial'; ctx.textBaseline = 'top'; ctx.textAlign = 'start';
      ctx.fillText('مرحبا ABC 123 بالعالم', 550, 15);
      ctx.direction = 'ltr'; ctx.textAlign = 'left';
      const width = ctx.measureText('日本😀').width;
      ctx.save(); ctx.beginPath(); ctx.rect(10, 70, width - .1, 40); ctx.clip();
      ctx.fillText('日本😀SECRET', 10, 70); ctx.restore();
      return pdf.canvasLines(cv);
    } finally { cv.remove(); }
  })()`);
  assert.deepEqual(result[0].map(w => w.text), ['مرحبا', 'ABC', '123', 'بالعالم']);
  assert.ok(result[0][0].x > result[0][3].x);
  assert.ok(result[0][1].x < result[0][2].x);
  assert.equal(result[0][0].visual, 'ابحرم');
  assert.equal(result[1].map(w => w.text).join(''), '日本😀');
});

const setup = String.raw`(async () => {
  const M = L.model, wb = new M.Workbook({ font: { name: 'Arial', sz: 11 } });
  const sh = wb.addSheet('PDF cells');
  const put = (r, c, v, style = {}) => sh.put(r, c, { v, s: wb.styles.derive(0, style) });
  sh.cols[0] = { w: 28 }; sh.cols[1] = { w: 18 }; sh.cols[2] = { w: 16 };
  put(0, 0, 'First cell'); put(0, 1, 1234.5, { nf: '#,##0.00' });
  put(0, 2, 44927, { nf: 'yyyy-mm-dd' });
  put(1, 0, 'Wrapped words across several visible lines', { align: { wrap: true, v: 'top' } });
  sh.rows[1].ht = 60;
  put(2, 0, 'Visible prefix then this very long SECRETTAIL should be clipped');
  put(2, 1, 'Block overflow');
  put(3, 0, 'Overflow into empty neighbours');
  put(4, 0, 'Turned needle', { align: { rot: 45, v: 'center' } });
  put(4, 1, 'Vertical', { align: { rot: 90, v: 'center' } }); sh.rows[4].ht = 90;
  put(5, 0, 'Shrink this entire phrase to fit', { align: { shrink: true } });
  put(6, 0, '日本語 😀'); put(6, 1, 'مرحبا بالعالم');
  put(7, 0, 'HIDDENROW'); sh.rows[7].hidden = true;
  put(8, 0, 'FORMATSECRET', { nf: ';;;' }); put(8, 2, 'Last cell');
  sh.print.header = '&L&12Ledger PDF header'; sh.print.footer = '&CPage &P of &N';
  wb.props.title = 'Canvas text checks';
  L.app.addBook(wb, 'PDF cells.xlsx');
  await document.fonts.ready;
})()`;

async function exported(opts = {}) {
  const result = await page.evaluate(`(async () => {
    const build = L.buildPDF, save = L.saveFile;
    let pages, blob;
    L.buildPDF = (list, info) => { pages = list.map(({ jpeg, ...p }) => p); return build(list, info); };
    L.saveFile = async (name, value) => { blob = value; return 'saved'; };
    try {
      if (!await L.print.exportPDF({ quiet: true, ...${JSON.stringify(opts)} })) throw Error('PDF export failed');
      return { pages, data: (await L.blobToDataURL(blob)).split(',')[1] };
    } finally { L.buildPDF = build; L.saveFile = save; }
  })()`);
  const file = path.join(dir, 'cells.pdf');
  fs.writeFileSync(file, Buffer.from(result.data, 'base64'));
  return { ...result, file, text: execFileSync('pdftotext', ['-raw', file, '-'], { encoding: 'utf8' }) };
}

test('Ledger PDF records displayed cell values, wraps, overflow, shrink, rotation and headers/footers', async () => {
  await page.evaluate(setup);
  const result = await exported();
  assert.equal(result.pages.length, 1);
  assert.equal(result.pages[0].pw, 612);
  assert.equal(result.pages[0].ph, 792);
  for (const text of ['First cell', '1,234.50', '2023-01-01', 'Wrapped words', 'visible lines',
    'Overflow into empty neighbours', 'Shrink this entire phrase to fit', '日本語', '😀', 'Last cell',
    'Ledger PDF header', 'Page 1 of 1']) assert.ok(result.text.replace(/\s+/gu, ' ').includes(text), text);
  assert.doesNotMatch(result.text, /HIDDENROW|FORMATSECRET|SECRETTAIL/);
  const copied = execFileSync('python3', ['-I', '-c',
    'import pypdfium2 as p,sys; d=p.PdfDocument(sys.argv[1]); print(d[0].get_textpage().get_text_range())',
    result.file], { encoding: 'utf8' });
  assert.match(copied, /Turned needle/);
  assert.match(copied, /Vertical/);
  assert.ok(result.pages[0].lines.flat().some(w => Math.abs(w.angle + 45) < .01));
  assert.deepEqual(page.errors, []);
  assert.deepEqual(page.requests, []);
});

test('repeated print rows and columns copy in row order across their separate canvases', async () => {
  await page.evaluate(`(async () => {
    const wb = new L.model.Workbook({ font: { name: 'Arial', sz: 11 } }), sh = wb.addSheet('Titles');
    for (let r = 0; r < 4; r++) for (let c = 0; c < 4; c++) sh.put(r, c, { v: 'R' + (r + 1) + 'C' + (c + 1) });
    wb.names.push({ name: '_xlnm.Print_Titles', scope: 0, ref: 'Titles!$1:$1,Titles!$A:$A' });
    sh.print.rowBreaks = [2]; sh.print.colBreaks = [2];
    sh.print.footer = '&CPage &P of &N';
    L.app.addBook(wb, 'Titles.xlsx'); await document.fonts.ready;
  })()`);
  const result = await exported({ from: 4, to: 4 });
  assert.equal(result.pages.length, 1);
  const words = result.pages[0].lines?.flat().map(w => w.text) || [];
  assert.deepEqual(words.filter(text => /^R\dC\d$/.test(text)),
    ['R1C1', 'R1C3', 'R1C4', 'R3C1', 'R3C3', 'R3C4', 'R4C1', 'R4C3', 'R4C4']);
  assert.match(result.text, /Page 4 of 4/);
});

test('worksheet and chart-sheet PDFs keep SVG labels once and leave the page picture unchanged', async () => {
  for (const kind of ['chart', 'shape', 'chartsheet']) {
    const result = await page.evaluate(`(async () => {
      const pdf = await L.pdfModule(true), wb = new L.model.Workbook(), sh = wb.addSheet('Chart labels');
      sh.put(0, 0, { v: 'Cell label' });
      const old = { render: L.xchart.render, element: L.drawing.element, update: L.drawing.update,
        build: L.buildPDF, save: L.saveFile, record: pdf.recordCanvas, lines: pdf.canvasLines };
      function svg() {
        const node = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
        node.setAttribute('width', '400'); node.setAttribute('height', '100');
        const label = document.createElementNS(node.namespaceURI, 'text');
        label.setAttribute('x', '10'); label.setAttribute('y', '45');
        label.setAttribute('font-family', 'Carlito'); label.setAttribute('font-size', '24');
        label.textContent = 'Printed SVG label'; node.append(label);
        return node;
      }
      try {
        L.xchart.render = svg;
        L.drawing.element = () => { const e = document.createElement('div'); e.append(svg()); return e; };
        L.drawing.update = () => {};
        sh.drawings = [{ kind: ${JSON.stringify(kind === 'chartsheet' ? 'chart' : kind)}, chart: {},
          anchor: { from: { r: 2, c: 0 }, w: 300, h: 75 } }];
        if (${JSON.stringify(kind)} === 'chartsheet') sh.kind = 'chartsheet';
        L.app.addBook(wb, 'Labels.xlsx'); await document.fonts.load('24px Carlito');
        let pages;
        L.buildPDF = list => { pages = list; return new Blob(); };
        L.saveFile = async () => 'saved';
        await L.print.exportPDF({ quiet: true });
        const recorded = pages;
        pdf.recordCanvas = ctx => ctx; pdf.canvasLines = () => [];
        await L.print.exportPDF({ quiet: true });
        return { words: recorded[0].lines.flat().map(w => w.text),
          same: recorded[0].jpeg.every((b, i) => b === pages[0].jpeg[i]) };
      } finally {
        L.xchart.render = old.render; L.drawing.element = old.element; L.drawing.update = old.update;
        L.buildPDF = old.build; L.saveFile = old.save;
        pdf.recordCanvas = old.record; pdf.canvasLines = old.lines;
      }
    })()`);
    assert.equal(result.same, true, kind);
    for (const word of ['Printed', 'SVG', 'label']) {
      assert.equal(result.words.filter(w => w === word).length, word === 'label' && kind !== 'chartsheet' ? 2 : 1,
        kind + ': ' + word);
    }
  }
});
