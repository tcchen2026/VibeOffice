// PDF page pictures use the same vendored fonts as the page, at 300 dpi (40 megapixels at most).
// These pixel probes run through the actual SVG/image boundary: a page's @font-face alone does not cross it.
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../tools/ooxml/cdp.mjs';

let stop;
before(async () => { stop = await ensureServer(); });
after(() => stop?.());

for (const app of ['quire', 'lectern']) {
  test(app + ': a rasterized page keeps Carlito regular, bold and italic', async () => {
    const page = await openPage(app, { persistence: false, automated: true });
    try {
      const result = await page.evaluate(`(async () => {
        const el = document.createElement('div');
        el.style.cssText = 'width:420px;height:70px;background:white;color:black;line-height:60px';
        el.textContent = 'Minimum office WWW iii 123';
        document.body.append(el);
        const rows = [];
        try {
          for (const face of ['', 'bold ', 'italic ', 'bold italic ']) {
            const draw = async family => {
              el.style.font = face + '30px/60px ' + family + ',serif';
              await document.fonts.load(face + '30px Carlito');
              const canvas = ${app === 'quire'
                ? "await L.app.rasterizePage(el, 1, '', new Map())"
                : 'await L.app.rasterize(el, 420, 70, 1, { map: new Map() })'};
              return canvas.toDataURL();
            };
            const real = await draw('Carlito'), missing = await draw('PDFMissingFont');
            rows.push({ face, different: real !== missing });
          }
          return rows;
        } finally { el.remove(); }
      })()`);
      for (const row of result) assert.ok(row.different, app + ': ' + (row.face || 'regular') + ' uses Carlito');
      assert.deepEqual(page.errors, []);
      assert.deepEqual(page.requests, []);
    } finally { await page.close(); }
  });
}

test('PDF batches keep each page’s SVG paint references local to that page', async () => {
  const page = await openPage('quire', { persistence: false, automated: true });
  try {
    const same = await page.evaluate(`(async () => {
      const pdf = await L.pdfModule(), fonts = await L.embedFonts('Carlito');
      const pages = ['red', 'green', 'blue'].map(color => ({ w: 80, h: 80, scale: 1,
        svg: '<svg xmlns="http://www.w3.org/2000/svg" width="80" height="80">' +
          '<defs><linearGradient id="paint"><stop stop-color="' + color + '"/>' +
          '<stop offset="1" stop-color="white"/></linearGradient></defs>' +
          '<rect width="80" height="80" fill="url(#paint)"/></svg>' }));
      const expected = [];
      for (const p of pages) {
        const c = document.createElement('canvas'); c.width = c.height = 80;
        const image = await L.loadImage('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(p.svg));
        c.getContext('2d').drawImage(image, 0, 0);
        expected.push(c.toDataURL());
      }
      const same = [];
      for await (const { canvas, index } of pdf.pictures(pages, p => p, fonts)) {
        same.push(canvas.toDataURL() === expected[index]);
      }
      return same;
    })()`);
    assert.deepEqual(same, [true, true, true]);
  } finally { await page.close(); }
});

for (const kind of ['chart', 'shape', 'chartsheet']) {
  test('Ledger: ' + kind + ' text keeps its web font inside the SVG image', async () => {
    const page = await openPage('ledger', { persistence: false, automated: true });
    try {
      const result = await page.evaluate(`(async () => {
        const sh = L.grid.sheet(), P = L.print;
        let family = 'Carlito';
        const svg = () => {
          const node = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          node.setAttribute('width', '400'); node.setAttribute('height', '100');
          const t = document.createElementNS(node.namespaceURI, 'text');
          t.setAttribute('x', '5'); t.setAttribute('y', '45');
          t.setAttribute('font-family', family + ',serif'); t.setAttribute('font-size', '30');
          t.setAttribute('font-weight', '700'); t.setAttribute('font-style', 'italic');
          t.textContent = 'Minimum office WWW iii 123'; node.append(t);
          return node;
        };
        L.xchart.render = svg;
        L.drawing.element = () => { const e = document.createElement('div'); e.append(svg()); return e; };
        L.drawing.update = () => {};
        sh.drawings = [{ kind: ${JSON.stringify(kind === 'chartsheet' ? 'chart' : kind)}, chart: {},
          anchor: { from: { r: 0, c: 0 }, w: 300, h: 75 } }];
        await document.fonts.load('bold italic 30px Carlito');
        const draw = async () => {
          if (${JSON.stringify(kind)} === 'chartsheet') {
            sh.kind = 'chartsheet';
            let jpeg;
            L.buildPDF = pages => { jpeg = Array.from(pages[0].jpeg); return new Blob(); };
            L.saveFile = async () => 'saved';
            if (!await P.exportPDF({ quiet: true })) throw Error('Export failed');
            return JSON.stringify(jpeg);
          }
          const pag = P.paginate(sh);
          const c = await P.renderPage(sh, { r1: 0, r2: 7, c1: 0, c2: 6 },
            { pag, page: 1, pages: 1 }, 1);
          return c.toDataURL();
        };
        const real = await draw(); family = 'PDFMissingFont'; const missing = await draw();
        return real !== missing;
      })()`);
      assert.ok(result, kind + ': embedded Carlito differs from the missing font');
      assert.deepEqual(page.errors, []);
      assert.deepEqual(page.requests, []);
    } finally { await page.close(); }
  });
}

test('PDF resolution: Letter and A3 stay at 300 dpi; A0 stays within 40 megapixels', async () => {
  const page = await openPage('quire', { persistence: false, automated: true });
  try {
    const result = await page.evaluate(`(() => {
      const scale = (w, h, k) => { const s = L.pdfScale(w, h, k); return [s, Math.round(w * s), Math.round(h * s)]; };
      return { dpi: L.PDF_DPI, letter: scale(612, 792, L.PDF_DPI / 72),
        a3: scale(842, 1191, L.PDF_DPI / 72), a0: scale(2384, 3370, L.PDF_DPI / 72),
        landscape: scale(3370, 2384, L.PDF_DPI / 72), thumb: scale(816, 1056, 400 / 816) };
    })()`);
    assert.equal(result.dpi, 300);
    assert.deepEqual(result.letter, [300 / 72, 2550, 3300]);
    assert.equal(result.a3[0], 300 / 72);
    for (const box of [result.a0, result.landscape]) {
      assert.ok(box[1] * box[2] <= 40_000_000);
      assert.ok(box[1] * box[2] > 39_970_000);
    }
    assert.equal(result.thumb[1], 400, 'the thumbnail keeps its own scale');
  } finally { await page.close(); }
});

for (const app of ['quire', 'lectern', 'ledger']) {
  test(app + ': PDF pages use 300 dpi and prepare font data once per export', async () => {
    const page = await openPage(app, { persistence: false, automated: true });
    try {
      const result = await page.evaluate(`(async () => {
        let calls = 0, pages;
        const embed = L.embedFonts;
        L.embedFonts = async (...args) => { calls++; return embed(...args); };
        L.buildPDF = list => { pages = list.map(p => ({ iw: p.iw, ih: p.ih, pw: p.pw, ph: p.ph })); return new Blob(); };
        L.saveFile = async () => 'saved';
        if (${JSON.stringify(app)} === 'quire') {
          const { D } = L, doc = D.newDoc();
          doc.main.blocks = [D.para([D.text('Calibri font on page one.', { font: 'Calibri' })]),
            D.para([D.text('Calibri font on page two.', { font: 'Calibri' })], { pageBreakBefore: true })];
          L.app.loadDoc(doc, 'PDF fonts');
          await document.fonts.ready;
          await L.app.buildPDF();
        } else if (${JSON.stringify(app)} === 'lectern') {
          L.pres.slides.push(structuredClone(L.pres.slides[0]));
          await L.app.buildPDF('slides');
        } else {
          await L.app.openFile(new File(['Font,Value\\nCalibri,42'], 'PDF.csv'));
          await document.fonts.ready;
          await L.print.exportPDF({ quiet: true });
        }
        L.ui.busy(false);
        return { calls, pages };
      })()`);
      assert.ok(result.pages.length, 'pages produced');
      if (app !== 'ledger') assert.equal(result.pages.length, 2, 'two pages share the prepared font data');
      for (const p of result.pages) {
        assert.ok(Math.abs(p.iw - p.pw * 300 / 72) <= 2, 'horizontal resolution: ' + JSON.stringify(p));
        assert.ok(Math.abs(p.ih - p.ph * 300 / 72) <= 2, 'vertical resolution: ' + JSON.stringify(p));
      }
      assert.equal(result.calls, app === 'ledger' ? 0 : 1, 'font data prepared once; cells need no SVG fonts');
      assert.deepEqual(page.errors, []);
      assert.deepEqual(page.requests, []);
    } finally { await page.close(); }
  });
}

test('PDF export can encode a page without waiting for an idle toBlob callback', async () => {
  const page = await openPage('quire', { persistence: false, automated: true });
  try {
    const result = await page.evaluate(`(async () => {
      const native = HTMLCanvasElement.prototype.toBlob;
      HTMLCanvasElement.prototype.toBlob = function (callback, type, quality) {
        if (type === 'image/jpeg') throw new Error('The browser deferred its asynchronous JPEG encoder.');
        return native.call(this, callback, type, quality);
      };
      try {
        const pdf = await L.app.buildPDF();
        return { type: pdf.type, size: pdf.size };
      } finally {
        HTMLCanvasElement.prototype.toBlob = native;
        L.ui.busy(false);
      }
    })()`);
    assert.equal(result.type, 'application/pdf');
    assert.ok(result.size > 1000);
  } finally { await page.close(); }
});

test('Lectern: portrait and landscape preview pages use 300 dpi', async () => {
  const page = await openPage('lectern', { persistence: false, automated: true });
  try {
    const pages = await page.evaluate(`(async () => {
      const scroll = document.createElement('div');
      for (const [w, h] of [[816, 1056], [1056, 816]]) {
        const el = document.createElement('div');
        el.className = 'pv-page';
        el.style.cssText = 'width:' + w + 'px;height:' + h + 'px;font:20px Carlito';
        el.textContent = 'Preview uses the same font and resolution as slides.';
        scroll.append(el);
      }
      document.body.append(scroll);
      const previous = L.panes.preview.scroll;
      L.panes.preview.scroll = scroll;
      let result;
      L.buildPDF = pages => { result = pages.map(p => [p.iw, p.ih, p.pw, p.ph]); return new Blob(); };
      try { await L.app.buildPDF('preview'); return result; }
      finally { L.panes.preview.scroll = previous; scroll.remove(); L.ui.busy(false); }
    })()`);
    assert.deepEqual(pages, [[2550, 3300, 612, 792], [3300, 2550, 792, 612]]);
    assert.deepEqual(page.errors, []);
  } finally { await page.close(); }
});

for (const app of ['quire', 'lectern']) {
  test(app + ': distinct font-heavy pages share font parsing without changing their JPEG pixels', async () => {
    const page = await openPage(app, { persistence: false, automated: true });
    try {
      const result = await page.evaluate(`(async () => {
        const app = ${JSON.stringify(app)}, elements = [];
        for (let n = 0; n < 5; n++) {
          const el = document.createElement('div');
          el.style.cssText = 'width:794px;height:1123px;background:white;color:black;position:relative';
          el.innerHTML = '<p style="font:24px Carlito">Different page ' + n + '</p>' +
            '<p style="font:bold 24px Carlito">Bold</p><p style="font:italic 24px Carlito">Italic</p>' +
            '<p style="font:bold italic 24px Carlito">Both</p>';
          document.body.append(el);
          elements.push(el);
        }
        await document.fonts.ready;
        const image = L.loadImage, individual = [];
        let actual, loads = 0, maxPixels = 0;
        const encode = c => new Promise(resolve => c.toBlob(async blob => {
          resolve(Array.from(new Uint8Array(await blob.arrayBuffer())));
        }, 'image/jpeg', .92));
        try {
          if (app === 'quire') {
            L.layout.pages = elements.map(el => ({ el, sect: { pgW: 595.5, pgH: 842.25 } }));
          } else {
            L.pres.W = 794; L.pres.H = 1123;
            L.pres.slides = elements.map((el, n) => ({ id: 'PDF-test-' + n }));
            L.render.slide = (pres, slide) => elements[pres.slides.indexOf(slide)];
          }
          L.buildPDF = list => { actual = list.map(p => Array.from(p.jpeg)); return new Blob(); };
          // Capture the exact isolated SVGs used by the app, then compare the batched export with them.
          const fonts = await L.embedFonts(elements.map(el => el.outerHTML).join(''));
          const scale = app === 'quire' ? 300 / 96 : 300 / 72;
          for (const el of elements) {
            const c = app === 'quire' ? await L.app.rasterizePage(el, scale, '', new Map(), fonts)
              : await L.app.rasterize(el, 794, 1123, scale, { map: new Map(), css: ' ', fonts });
            individual.push(await encode(c));
            c.width = c.height = 1;
          }
          L.loadImage = async src => {
            if (src.startsWith('data:image/svg+xml')) {
              loads++;
              const svg = new DOMParser().parseFromString(decodeURIComponent(src.split(',')[1]), 'image/svg+xml');
              const el = svg.documentElement;
              maxPixels = Math.max(maxPixels,
                Math.round(+el.getAttribute('width') * scale) * Math.round(+el.getAttribute('height') * scale));
            }
            return image(src);
          };
          await L.app.buildPDF(app === 'lectern' ? 'slides' : undefined);
          return { loads, maxPixels, same: actual.map((bytes, n) =>
            bytes.length === individual[n].length && bytes.every((byte, i) => byte === individual[n][i])) };
        } finally {
          L.loadImage = image;
          elements.forEach(el => el.remove());
          L.ui.busy(false);
        }
      })()`);
      assert.ok(result.loads < 5, JSON.stringify(result));
      assert.ok(result.maxPixels <= 40_000_000, JSON.stringify(result));
      assert.deepEqual(result.same, [true, true, true, true, true]);
      assert.deepEqual(page.errors, []);
    } finally { await page.close(); }
  });
}
