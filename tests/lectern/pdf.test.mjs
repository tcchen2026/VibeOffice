// Lectern's PDF words come from attached slide/preview DOM; the page pictures stay unchanged.
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
  page = await openPage('lectern', { persistence: false, automated: true });
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'lectern-pdf-'));
});
after(async () => {
  if (page) await page.close();
  stop?.();
  if (dir) fs.rmSync(dir, { recursive: true, force: true });
});

test('PDF rotation keeps line order and word dimensions through nested thumbnail scaling', async () => {
  const results = await page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true), results = [];
    for (const [angle, outer, center] of [[45, 0, false], [90, 0, false], [-30, 0, false], [45, 20, true]]) {
      const el = document.createElement('div');
      el.style.cssText = 'position:fixed;left:0;top:0;width:480px;height:360px;font:16px/24px Arial';
      el.innerHTML = '<div style="position:absolute;left:30px;top:40px;transform:scale(.6);transform-origin:0 0">' +
        '<div class="turn" style="position:absolute;left:80px;top:50px;width:180px;height:100px;transform-origin:0 0">' +
        '<span>Alpha</span> <b>beta</b><br>Next row</div></div>';
      document.body.append(el);
      try {
        const turn = el.querySelector('.turn'), span = turn.querySelector('span');
        const raw = span.getBoundingClientRect(), base = turn.getBoundingClientRect();
        const local = { x: (raw.left - base.left) / .6, y: (raw.top - base.top) / .6 };
        const matrix = new DOMMatrix().translate(30, 40).scale(.6).rotate(outer).translate(80, 50);
        if (center) matrix.translateSelf(90, 50);
        matrix.rotateSelf(angle);
        if (center) matrix.translateSelf(-90, -50);
        const expected = matrix.transformPoint(local);
        turn.parentElement.style.transform = 'scale(.6) rotate(' + outer + 'deg)';
        if (center) turn.style.transformOrigin = '50% 50%';
        turn.style.transform = 'rotate(' + angle + 'deg)';
        const before = el.outerHTML;
        const lines = pdf.collectDOM(el, 1, { transforms: true });
        results.push({ angle: angle + outer, lines,
          expected: { x: expected.x, y: expected.y, w: raw.width, h: raw.height },
          unchanged: before === el.outerHTML });
      } finally { el.remove(); }
    }
    return results;
  })()`);
  for (const result of results) {
    assert.deepEqual(result.lines.map(line => line.map(w => w.text).join(' ')), ['Alpha beta', 'Next row']);
    const word = result.lines[0][0];
    assert.ok(Math.abs(word.angle - result.angle) < 0.01, 'rotation ' + result.angle);
    for (const key of ['x', 'y', 'w', 'h']) {
      assert.ok(Math.abs(word[key] - result.expected[key]) < 0.15, key + ' at ' + result.angle);
    }
    assert.equal(result.unchanged, true, 'measuring leaves the live page unchanged');
  }
});

test('PDF rotation clips text in the rotated ancestor plane and at the page edge', async () => {
  const result = await page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true), el = document.createElement('div');
    el.style.cssText = 'position:fixed;left:0;top:0;width:300px;height:240px;font:16px/24px monospace';
    el.innerHTML = '<div style="position:absolute;left:100px;top:50px;width:62px;height:24px;' +
      'overflow:hidden;white-space:nowrap;transform:rotate(45deg);transform-origin:0 0">' +
      'ABCDEFGHIJKLMNO<br>HIDDENROW</div>' +
      '<div style="position:absolute;left:350px;top:180px;transform:rotate(90deg)">OUTSIDE</div>';
    document.body.append(el);
    try { return pdf.collectDOM(el, 1, { transforms: true }); }
    finally { el.remove(); }
  })()`);
  const words = result.flat(), text = words.map(w => w.text).join(' ');
  assert.match(text, /ABCDEF/);
  assert.doesNotMatch(text, /IJKLMNO|HIDDENROW|OUTSIDE/);
  assert.ok(words.every(w => Math.abs(w.angle - 45) < 0.01));
});

test('rotated LettersArt keeps its transparent source text once and excludes repeated SVG slices', async () => {
  const result = await page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true), el = document.createElement('div');
    el.style.cssText = 'position:fixed;left:0;top:0;width:400px;height:300px;font:20px/28px Arial';
    el.innerHTML = '<div class="shp warped" style="position:absolute;left:100px;top:100px;' +
      'width:180px;height:60px;transform:rotate(30deg)">' +
      '<svg class="warp" aria-hidden="true"><text>Repeated slices Repeated slices</text></svg>' +
      '<div class="tx" style="position:absolute;left:0;top:0;opacity:0">Bent words</div></div>';
    document.body.append(el);
    try { return pdf.collectDOM(el, 1, { transforms: true }); }
    finally { el.remove(); }
  })()`);
  assert.equal(result.flat().map(w => w.text).join(' '), 'Bent words');
  assert.ok(result.flat().every(w => Math.abs(w.angle - 30) < 0.01));
});

const setup = String.raw`(async () => {
  const M = L.model, T = L.txt, pres = M.newPresentation({ empty: true, w: 720, h: 540 });
  const first = M.newSlide(pres, 'text'), hidden = M.newSlide(pres, 'text'), last = M.newSlide(pres, 'text');
  const text = (slide, title, body) => {
    T.setPlain(slide.shapes[0].tx, title);
    T.setPlain(slide.shapes[1].tx, body);
  };
  text(first, 'First title', 'Hello styled world\n日本語 😀');
  text(hidden, 'HIDDENTITLE', 'HIDDENBODY');
  hidden.hidden = true;
  text(last, 'Second title', 'Second body');
  const turned = M.newTextBox(pres, last, 250, 220, 200, 70, true);
  turned.rot = 45;
  turned.tx.ps = [T.para('Rotated needle', {}, { sz: 24 })];
  last.shapes.push(turned);
  first.notes = 'First speaker notes'; last.notes = 'Second speaker notes';
  pres.slides = [first, hidden, last];
  L.app.loadPres(pres, 'PDF test', { saved: true });
  await document.fonts.ready;
})()`;

async function exported(source = 'slides') {
  const result = await page.evaluate(`(async () => {
    const build = L.buildPDF, hosts = () => document.querySelectorAll('body > [inert]').length;
    const before = hosts();
    let pages;
    L.buildPDF = (list, info) => { pages = list.map(({ jpeg, ...p }) => p); return build(list, info); };
    try {
      const blob = await L.app.buildPDF(${JSON.stringify(source)});
      L.ui.busy(false);
      return { pages, data: (await L.blobToDataURL(blob)).split(',')[1],
        unchanged: before === hosts() };
    } finally { L.buildPDF = build; }
  })()`);
  const file = path.join(dir, source + '.pdf');
  fs.writeFileSync(file, Buffer.from(result.data, 'base64'));
  return { ...result, file, text: execFileSync('pdftotext', ['-raw', file, '-'], { encoding: 'utf8' }) };
}

test('Lectern slide PDFs collect detached slide text, omit hidden slides and keep slide dimensions', async () => {
  await page.evaluate(setup);
  const result = await exported();
  assert.equal(result.pages.length, 2);
  assert.ok(result.pages.every(p => p.pw === 720 && p.ph === 540));
  for (const text of ['First title', 'Hello styled world', '日本語', '😀', 'Second title']) {
    assert.ok(result.text.includes(text), text);
  }
  // Poppler's raw extractor separates diagonal glyphs with newlines; PDFium preserves the phrase.
  assert.ok(result.text.replace(/\s/gu, '').includes('Rotatedneedle'));
  const copied = execFileSync('python3', ['-I', '-c',
    'import pypdfium2 as p,sys; d=p.PdfDocument(sys.argv[1]); print(d[1].get_textpage().get_text_range())',
    result.file], { encoding: 'utf8' });
  assert.match(copied, /Rotated needle/);
  assert.doesNotMatch(result.text, /HIDDENTITLE|HIDDENBODY|speaker notes/);
  const rotated = result.pages[1].lines.flat().filter(w => ['Rotated', 'needle'].includes(w.text));
  assert.equal(rotated.length, 2);
  assert.ok(rotated.every(w => Math.abs(w.angle - 45) < 0.01));
  assert.equal(result.unchanged, true);
  assert.deepEqual(page.errors, []);
  assert.deepEqual(page.requests, []);
});

test('Lectern handouts, notes and outline PDFs retain preview text, labels and paper size', async () => {
  await page.evaluate(setup);
  for (const what of ['handouts', 'notes', 'outline']) {
    await page.evaluate(`(() => {
      Object.assign(L.app.printOpts, { what: ${JSON.stringify(what)}, per: 2, hidden: false, range: null });
      L.app.setView('preview');
      L.panes.preview.render();
    })()`);
    const result = await exported('preview');
    assert.ok(result.pages.every(p => p.pw === 612 && p.ph === 792));
    assert.match(result.text, /First title/);
    assert.match(result.text, /Second title/);
    assert.doesNotMatch(result.text, /HIDDENTITLE|HIDDENBODY/);
    if (what === 'notes') {
      assert.match(result.text, /First speaker notes/);
      assert.match(result.text, /Second speaker notes/);
    } else assert.doesNotMatch(result.text, /speaker notes/);
    assert.equal(result.unchanged, true);
  }
});

test('Lectern releases a temporary measuring host when collection fails', async () => {
  await page.evaluate(setup);
  const result = await page.evaluate(`(async () => {
    const pdf = await L.pdfModule(true), collect = pdf.collectDOM;
    const hosts = () => document.querySelectorAll('body > [inert]').length, before = hosts();
    pdf.collectDOM = () => { throw new Error('PDF test collection failure'); };
    let error = '';
    try { await L.app.buildPDF('slides'); } catch (e) { error = e.message; }
    finally { pdf.collectDOM = collect; L.ui.busy(false); }
    return { error, unchanged: before === hosts() };
  })()`);
  assert.equal(result.error, 'PDF test collection failure');
  assert.equal(result.unchanged, true);
});
