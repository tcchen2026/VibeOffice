// PDF navigation uses the exported pages, including ranges, hidden slides and repeated print titles.
// Real app rendering supplies the geometry; the shared writer's own tests check PDF object encoding.
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../tools/ooxml/cdp.mjs';

const apps = new Map();
let stop;
before(async () => {
  stop = await ensureServer();
  for (const app of ['quire', 'lectern', 'ledger']) {
    const page = await openPage(app, { persistence: false, automated: true });
    await page.send('Emulation.setFocusEmulationEnabled', { enabled: true });
    apps.set(app, page);
  }
});
after(async () => {
  for (const page of apps.values()) await page.close();
  stop?.();
});

async function exported(app, arg) {
  const page = apps.get(app);
  const result = await page.evaluate(`(async () => {
    const build = L.buildPDF, save = L.saveFile;
    let result;
    L.buildPDF = (pages, info) => {
      result = { pages: pages.map(({ jpeg, lines, ...p }) => p), info };
      return build(pages, { ...info, date: '2020-01-02T03:04:05Z' });
    };
    L.saveFile = async () => 'saved';
    try {
      ${app === 'ledger' ? 'await L.print.exportPDF({ quiet: true, ...' + JSON.stringify(arg || {}) + ' });' :
        'await L.app.buildPDF(' + JSON.stringify(arg) + ');'}
      return result;
    } finally { L.buildPDF = build; L.saveFile = save; L.ui.busy(false); }
  })()`);
  assert.ok(result, app + ' exported');
  assert.deepEqual(page.errors, [], app + ' browser errors');
  assert.deepEqual(page.requests, [], app + ' external requests');
  return result;
}

test('PDF link rectangles wrap, clip and exclude hidden content without changing the DOM', async () => {
  const result = await apps.get('quire').evaluate(`(async () => {
    const pdf = await L.pdfModule(true), el = document.createElement('div');
    el.style.cssText = 'position:fixed;left:-10000px;top:0;width:240px;height:240px;font:16px/24px monospace';
    el.innerHTML = '<div style="width:100px"><span data-link="wrapped">Alpha beta gamma delta</span></div>' +
      '<div style="width:50px;height:24px;overflow:hidden"><span data-link="clipped">ABCDEFGHIJKLM</span></div>' +
      '<div style="display:none" data-link="display">Hidden</div>' +
      '<span style="visibility:hidden" data-link="visibility">Hidden</span>' +
      '<div style="opacity:0" data-link="opacity">Hidden</div>' +
      '<div aria-hidden="true" data-link="aria">Decoration</div>' +
      '<div style="position:absolute;top:260px" data-link="outside">Outside</div>';
    document.body.append(el);
    try {
      const before = el.outerHTML;
      const links = pdf.collectLinks(el, .75, node => ({ url: node.dataset.link }));
      return { links, unchanged: before === el.outerHTML };
    } finally { el.remove(); }
  })()`);
  assert.equal(result.unchanged, true);
  assert.ok(result.links.filter(l => l.url === 'wrapped').length > 1, 'one link area per wrapped line');
  assert.deepEqual([...new Set(result.links.map(l => l.url))], ['wrapped', 'clipped']);
  const clipped = result.links.find(l => l.url === 'clipped');
  assert.ok(clipped.w > 36 && clipped.x + clipped.w <= 37.51);
});

test('PDF link rectangles follow rotated and clipped shapes through nested thumbnail scaling', async () => {
  const result = await apps.get('lectern').evaluate(`(async () => {
    const pdf = await L.pdfModule(true), el = document.createElement('div');
    el.style.cssText = 'position:fixed;left:-10000px;top:0;width:300px;height:240px';
    el.innerHTML = '<div style="position:absolute;left:80px;top:60px;transform:scale(.6);transform-origin:0 0">' +
      '<div data-link="shape" style="width:180px;height:60px;transform:rotate(30deg);transform-origin:0 0"></div>' +
      '</div><div style="position:absolute;left:280px;top:200px;width:80px;height:80px;' +
      'overflow:hidden;transform:rotate(30deg);transform-origin:0 0">' +
      '<div data-link="edge" style="width:140px;height:100px"></div></div>';
    document.body.append(el);
    try {
      const origin = el.getBoundingClientRect(), r = el.querySelector('[data-link]').getBoundingClientRect();
      const before = el.outerHTML;
      return { links: pdf.collectLinks(el, 1, node => ({ url: node.dataset.link }), { transforms: true }),
        expected: { x: r.left - origin.left, y: r.top - origin.top, w: r.width, h: r.height },
        unchanged: before === el.outerHTML };
    } finally { el.remove(); }
  })()`);
  assert.equal(result.unchanged, true);
  assert.equal(result.links.length, 2);
  for (const key of ['x', 'y', 'w', 'h']) {
    assert.ok(Math.abs(result.links[0][key] - result.expected[key]) < .1, key);
  }
  assert.ok(result.links[1].x + result.links[1].w <= 300.01);
  assert.ok(result.links[1].y + result.links[1].h <= 240.01);
});

const quireSetup = String.raw`(async () => {
  const D = L.D, doc = D.newDoc();
  doc.props.language = 'fr-CA';
  doc.main.blocks = [
    D.para([D.text('First heading')], { outline: 0 }),
    D.para([D.text('Go to the last heading', { link: '#last' })]),
    D.para([D.text('External link', { link: 'https://example.test/日本語%20space?q=(x)' })]),
    D.para([D.text('Hidden destination', { link: '#missing' })]),
    D.para([D.text('Child heading')], { outline: 1, pageBreakBefore: true }),
    D.para([D.text('Back to first', { link: '#first' })]),
    D.para([D.item('bs', { id: '2', name: 'last' }), D.text('First heading'), D.item('be', { id: '2' })],
      { outline: 0, pageBreakBefore: true }),
  ];
  doc.main.blocks[0].runs.unshift(D.item('bs', { id: '1', name: 'first' }));
  doc.main.blocks[0].runs.push(D.item('be', { id: '1' }));
  L.app.loadDoc(doc, 'PDF links.docx', { saved: true });
  await document.fonts.ready;
  L.layout.render();
})()`;

test('Quire PDF links, nested heading bookmarks and document language follow actual pages', async () => {
  await apps.get('quire').evaluate(quireSetup);
  const result = await exported('quire');
  assert.equal(result.pages.length, 3);
  assert.equal(result.info.lang, 'fr-CA');
  assert.deepEqual(result.info.outline.map(o => [o.title, o.page, o.level]),
    [['First heading', 0, 1], ['Child heading', 1, 2], ['First heading', 2, 1]]);
  assert.ok(result.pages[0].links.some(l => l.page === 2 && l.top > 0));
  assert.ok(result.pages[0].links.some(l => l.url === 'https://example.test/日本語%20space?q=(x)'));
  assert.equal(result.pages[0].links.length, 2, 'unresolved destination omitted');
  assert.ok(result.pages[1].links.some(l => l.page === 0));
});

test('Quire partial PDF remaps destinations and omits links and headings on excluded pages', async () => {
  await apps.get('quire').evaluate(quireSetup);
  const result = await exported('quire', [1, 3]);
  assert.equal(result.pages.length, 2);
  assert.deepEqual(result.info.outline.map(o => [o.title, o.page]), [['First heading', 0], ['First heading', 1]]);
  assert.ok(result.pages[0].links.some(l => l.page === 1));
  const only = await exported('quire', [2]);
  assert.deepEqual(only.info.outline.map(o => [o.title, o.page]), [['Child heading', 0]]);
  assert.equal(only.pages[0].links.length, 0);
});

const lecternSetup = String.raw`(async () => {
  const M = L.model, T = L.txt, pres = M.newPresentation({ empty: true, w: 720, h: 540 });
  const first = M.newSlide(pres, 'text'), hidden = M.newSlide(pres, 'text'), last = M.newSlide(pres, 'text');
  pres.slides = [first, hidden, last]; hidden.hidden = true; pres.lang = 'ja-JP';
  for (const slide of pres.slides) {
    T.setPlain(slide.shapes[0].tx, slide === hidden ? 'Hidden title' : 'Duplicate title');
    T.setPlain(slide.shapes[1].tx, 'Linked words');
  }
  first.shapes[1].tx.ps[0].rs[0].link = { slide: last.id };
  last.shapes[1].tx.ps[0].rs[0].link = { url: 'https://example.test/slides' };
  const shape = M.newTextBox(pres, first, 100, 270, 200, 80, true);
  shape.rot = 30; shape.fill = { type: 'solid', color: '#00AA99' };
  shape.link = { slide: last.id }; first.shapes.push(shape);
  const missing = M.newTextBox(pres, first, 400, 270, 200, 80, true);
  missing.link = { slide: hidden.id }; first.shapes.push(missing);
  L.app.loadPres(pres, 'PDF links.pptx', { saved: true });
  await document.fonts.ready;
})()`;

test('Lectern PDF keeps text and rotated shape links, slide titles and language, excluding hidden targets', async () => {
  await apps.get('lectern').evaluate(lecternSetup);
  const result = await exported('lectern');
  assert.equal(result.pages.length, 2);
  assert.equal(result.info.lang, 'ja-JP');
  assert.deepEqual(result.info.outline.map(o => [o.title, o.page]),
    [['Duplicate title', 0], ['Duplicate title', 1]]);
  assert.equal(result.pages[0].links.length, 2);
  assert.ok(result.pages[0].links.every(l => l.page === 1));
  assert.ok(result.pages[1].links.some(l => l.url === 'https://example.test/slides'));
  const shape = result.pages[0].links.find(l => l.w > 200);
  assert.ok(shape && shape.h > 160, 'rotation changes the shape link rectangle');
});

test('Lectern handout, notes and outline PDF bookmarks use preview slots and filtered slide order', async () => {
  await apps.get('lectern').evaluate(lecternSetup);
  for (const what of ['handouts', 'notes', 'outline']) {
    await apps.get('lectern').evaluate(`(async () => {
      Object.assign(L.app.printOpts, { what: ${JSON.stringify(what)}, per: 2, hidden: false, range: null });
      L.app.setView('preview');
      await new Promise(requestAnimationFrame);
      L.panes.preview.render();
    })()`);
    const result = await exported('lectern', 'preview');
    assert.equal(result.info.outline.length, 2, what);
    assert.deepEqual(result.info.outline.map(o => o.page), what === 'notes' ? [0, 1] : [0, 0], what);
    assert.ok(result.info.outline[1].top > 0, what + ' destination inside paper');
    if (what !== 'outline') {
      const jump = result.pages[0].links.find(l => l.page !== undefined);
      assert.ok(jump, what + ' slide jump');
      assert.equal(jump.page, what === 'notes' ? 1 : 0);
      assert.equal(jump.top, result.info.outline[1].top);
    }
  }
});

const ledgerSetup = String.raw`(async () => {
  const wb = new L.model.Workbook(), sh = wb.addSheet('Titles'), other = wb.addSheet('Other sheet');
  for (let r = 0; r < 4; r++) for (let c = 0; c < 4; c++) sh.put(r, c, { v: 'R' + r + 'C' + c });
  other.put(0, 0, { v: 'Target cell' });
  wb.names.push({ name: '_xlnm.Print_Titles', scope: 0, ref: 'Titles!$1:$1,Titles!$A:$A' });
  wb.names.push({ name: 'Target', ref: "'Other sheet'!$A$1" });
  sh.print.rowBreaks = [2]; sh.print.colBreaks = [2];
  sh.links = [
    { ref: { r1: 0, r2: 0, c1: 0, c2: 0 }, target: 'https://example.test/repeated' },
    { ref: { r1: 2, r2: 2, c1: 2, c2: 2 }, location: 'Target' },
    { ref: { r1: 3, r2: 3, c1: 3, c2: 3 }, location: 'MissingName' },
  ];
  L.app.addBook(wb, 'PDF links.xlsx'); await document.fonts.ready;
})()`;

test('Ledger PDF repeats title-cell links and resolves named ranges without changing the active sheet', async () => {
  const page = apps.get('ledger');
  await page.evaluate(ledgerSetup);
  const before = await page.evaluate('[L.grid.wb.active, L.grid.sel().r, L.grid.sel().c]');
  const result = await exported('ledger', { what: 'workbook' });
  assert.equal(result.pages.length, 5);
  assert.deepEqual(result.info.outline.map(o => [o.title, o.page]), [['Titles', 0], ['Other sheet', 4]]);
  for (const pg of result.pages.slice(0, 4)) {
    assert.ok(pg.links.some(l => l.url === 'https://example.test/repeated'), 'repeated title link');
  }
  assert.ok(result.pages[3].links.some(l => l.page === 4 && l.top > 0));
  assert.equal(result.pages[3].links.length, 2, 'missing name omitted');
  assert.deepEqual(await page.evaluate('[L.grid.wb.active, L.grid.sel().r, L.grid.sel().c]'), before);
});

test('Ledger partial PDF remaps sheet bookmarks and only links to cells included in the output', async () => {
  await apps.get('ledger').evaluate(ledgerSetup);
  const result = await exported('ledger', { what: 'workbook', from: 4, to: 5 });
  assert.deepEqual(result.info.outline.map(o => [o.title, o.page]), [['Titles', 0], ['Other sheet', 1]]);
  assert.ok(result.pages[0].links.some(l => l.page === 1));
  const only = await exported('ledger', { from: 4, to: 4 });
  assert.deepEqual(only.info.outline.map(o => [o.title, o.page]), [['Titles', 0]]);
  assert.equal(only.pages[0].links.length, 1);
  assert.equal(only.pages[0].links[0].url, 'https://example.test/repeated');
});

test('Ledger PDF keeps HYPERLINK formula targets, merged links, hidden cells and centred print geometry', async () => {
  await apps.get('ledger').evaluate(`(async () => {
    const wb = new L.model.Workbook(), sh = wb.addSheet('Links');
    sh.put(0, 0, { f: 'HYPERLINK("https://"&B1,"Web link")', v: 'Web link' });
    sh.put(0, 1, { v: 'example.test/formula' });
    sh.put(1, 0, { f: 'HYPERLINK("#A5","Local link")', v: 'Local link' });
    sh.put(2, 0, { v: 'Merged link' });
    sh.merges.push({ r1: 2, r2: 2, c1: 0, c2: 1 });
    sh.put(3, 0, { v: 'Hidden link' }); sh.rows[3].hidden = true;
    sh.put(4, 0, { v: 'Destination' });
    sh.links = [
      { ref: { r1: 2, r2: 2, c1: 0, c2: 0 }, target: 'https://example.test/merged' },
      { ref: { r1: 3, r2: 3, c1: 0, c2: 0 }, target: 'https://example.test/hidden' },
    ];
    sh.print.hCenter = true; sh.print.vCenter = true; sh.print.scale = 75;
    L.app.addBook(wb, 'Links.xlsx'); await document.fonts.ready;
  })()`);
  const result = await exported('ledger');
  const links = result.pages[0].links;
  assert.equal(links.length, 3);
  const web = links.find(l => l.url === 'https://example.test/formula');
  const merged = links.find(l => l.url === 'https://example.test/merged');
  assert.ok(web && merged && merged.w > web.w * 1.9);
  assert.ok(links.some(l => l.page === 0 && l.top > merged.y));
  assert.ok(merged.x > 200 && merged.y > 300, 'print centring and scale applied');
});

test('Quire PDF resolves a bookmark inside a split paragraph to its fragment and text offset', async () => {
  const point = await apps.get('quire').evaluate(`(async () => {
    const D = L.D, doc = D.newDoc(), runs = [];
    doc.sect.pgW = 360; doc.sect.pgH = 360;
    for (let i = 0; i < 45; i++) {
      if (i === 40) runs.push(D.item('bs', { id: '8', name: 'Late' }));
      runs.push(D.text('Paragraph line ' + i), D.item('br'));
    }
    runs.push(D.item('be', { id: '8' }));
    doc.main.blocks = [D.para([D.text('Go to late text', { link: '#Late' })]), D.para(runs)];
    L.app.loadDoc(doc, 'Split bookmark.docx', { saved: true });
    await document.fonts.ready; L.layout.render();
    const b = L.fields.bookmarks().get('Late'), pg = L.layout.pageOf(b);
    const pt = L.layout.posToDom({ p: b.p, o: b.o + 1 }), range = document.createRange();
    range.setStart(pt.node, pt.offset); range.collapse(true);
    return { page: pg, top: (range.getClientRects()[0].top - L.layout.pages[pg].el.getBoundingClientRect().top)
      * .75 / L.layout.scale() };
  })()`);
  assert.ok(point.page > 1, 'bookmark is on a later paragraph fragment');
  const result = await exported('quire', [1, point.page + 1]);
  const link = result.pages[0].links[0];
  assert.equal(link.page, 1);
  assert.ok(Math.abs(link.top - point.top) < .1);
});

test('Lectern PDF converts static slide actions, clamps at deck ends, and omits show-history actions', async () => {
  await apps.get('lectern').evaluate(lecternSetup);
  await apps.get('lectern').evaluate(`(() => {
    const M = L.model, first = L.pres.slides[0]; first.shapes = [];
    for (const [i, action] of ['first', 'last', 'next', 'prev', 'lastViewed', 'end'].entries()) {
      const shape = M.newTextBox(L.pres, first, 20 + i * 100, 100, 80, 80, true);
      shape.link = { action }; first.shapes.push(shape);
    }
  })()`);
  const result = await exported('lectern');
  assert.deepEqual(result.pages[0].links.map(l => l.page), [0, 1, 1, 0]);
});

for (const app of ['quire', 'lectern', 'ledger']) {
  test(app + ': PDF omits hostile URI schemes and preserves web, mail, FTP, file and relative links', async () => {
    const blocked = ['javascript:alert(1)', 'vbscript:MsgBox(1)', 'data:text/html,<script>alert(1)</script>',
      ' \tJaVaScRiPt:alert(1)', 'java\tscript:alert(1)', 'vb\r\nscript:MsgBox(1)', ' \rDaTa:text/plain,blocked'];
    const kept = ['http://example.test/plain', 'https://example.test/file%20name?q=javascript:keep',
      'mailto:reader@example.test', 'ftp://example.test/file.txt', 'file:///C:/Documents/Report%20one.docx',
      '../data/report.xlsx#Sheet1!A1', './javascript:notes.txt', '/data:notes.txt', 'data/report.pdf',
      '//example.test/shared.pdf'];
    const page = apps.get(app);
    const result = await page.evaluate(`(async () => {
      const app = ${JSON.stringify(app)}, urls = ${JSON.stringify([...blocked, ...kept])};
      if (app === 'quire') {
        const D = L.D, doc = D.newDoc();
        doc.main.blocks = urls.map((url, i) => D.para([D.text('Link ' + i, { link: url })]));
        L.app.loadDoc(doc, 'Hostile links.docx', { saved: true });
      } else if (app === 'lectern') {
        const M = L.model, pres = M.newPresentation({ empty: true, w: 720, h: 540 });
        const slide = M.newSlide(pres, 'text');
        pres.slides = [slide];
        slide.shapes[1].tx.ps = urls.map((url, i) => L.txt.para('Link ' + i, {}, { sz: 12, link: { url } }));
        L.app.loadPres(pres, 'Hostile links.pptx', { saved: true });
      } else {
        const wb = new L.model.Workbook(), sh = wb.addSheet('Links');
        urls.forEach((url, r) => {
          sh.put(r, 0, { v: 'Link ' + r });
          sh.links.push({ ref: { r1: r, r2: r, c1: 0, c2: 0 }, target: url });
        });
        L.app.addBook(wb, 'Hostile links.xlsx');
      }
      await document.fonts.ready;
      if (app === 'quire') L.layout.render();
      const build = L.buildPDF, save = L.saveFile;
      let supplied = [], blob;
      L.buildPDF = (pages, info) => {
        supplied = pages.flatMap(p => (p.links || []).map(l => l.url).filter(Boolean));
        return build(pages, info);
      };
      L.saveFile = async (name, value) => { blob = value; return 'saved'; };
      try {
        if (app === 'ledger') await L.print.exportPDF({ quiet: true });
        else blob = await L.app.buildPDF();
        return { supplied, pdf: await blob.text() };
      } finally { L.buildPDF = build; L.saveFile = save; L.ui.busy(false); }
    })()`);
    for (const url of blocked) assert.ok(result.supplied.includes(url), 'hostile link reaches the shared writer');
    const uris = [...result.pdf.matchAll(/\/S \/URI \/URI \(((?:\\.|[^\\)])*)\)/g)].map(m => m[1]);
    assert.deepEqual(uris, kept);
    assert.deepEqual(page.errors, [], app + ' browser errors');
    assert.deepEqual(page.requests, [], app + ' external requests');
  });
}
