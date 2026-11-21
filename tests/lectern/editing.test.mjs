import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';
let page, stop;
before(async () => { stop = await ensureServer(); page = await openPage('lectern', { persistence: false, automated: true }); });
after(async () => { try { await page?.close(); } finally { stop?.(); } });
const run = async fn => {
  const result = await page.evaluate('(' + fn.toString() + ')()');
  assert.deepEqual(page.errors, []);
  return result;
};

test('F5, Rehearse Timings and a slide-show file start at the first visible slide; Shift+F5 can show a hidden slide', async () => {
  const result = await run(async () => {
    const p = L.model.newPresentation(); p.slides[0].hidden = true;
    p.slides.push(L.model.newSlide(p, 'blank'));
    L.app.loadPres(p, 'Show', { saved: true }); L.ed.goto(0);
    const calls = [], start = L.show.start;
    L.show.start = opts => { calls.push(opts?.from == null ? 'visible' : opts.from); };
    try {
      for (const cmd of ['showFromStart', 'rehearse', 'showFromCurrent']) L.ui.exec(cmd);
      const blob = await L.pptx.write(p); await __corpusHooks.open(new File([blob], 'Show.ppsx'));
    } finally { L.show.start = start; }
    return calls;
  });
  assert.deepEqual(result, ['visible', 'visible', 0, 'visible']);
});
test('Find and Replace All match across runs, preserve the first run link, and respect field and word boundaries', async () => {
  const result = await run(() => {
    const p = L.model.newPresentation(), slide = p.slides[0], sh = L.model.newTextBox(p, slide, 1, 1, 200, 100);
    slide.shapes = [sh];
    sh.tx.ps = [
      { rs: [{ t: 'Project ', b: true, link: { url: 'https://example.org/demo' } }, { t: 'Cedar' }, { t: ' by Project ' }, { t: 'Cedar' }] },
      { rs: [{ t: 'Project ' }, { t: 'Cedar', fld: 'customField' }] },
      { rs: [{ t: 'New' }, { t: 'Project ' }, { t: 'Cedar' }] },
    ].map(p => ({ lvl: 0, pp: {}, ...p }));
    L.app.loadPres(p, 'Project'); L.dlg.find(true);
    const dlg = document.querySelector('.dlg'), click = label => [...dlg.querySelectorAll('button')].find(b => b.textContent === label).click();
    dlg.querySelector('#fd-what').value = 'Project Cedar'; dlg.querySelector('#fd-with').value = 'Maple';
    dlg.querySelectorAll('input[type=checkbox]')[1].checked = true;
    click('Find Next'); const found = L.te.selectedText(); click('Replace All');
    [...document.querySelectorAll('.dlg button')].findLast(b => b.textContent === 'OK')?.click(); click('Close');
    const runs = sh.tx.ps[0].rs;
    return [found, sh.tx.ps.map(L.txt.paraText), runs[0].t, runs[0].b, runs[0].link?.url];
  });
  assert.deepEqual(result, ['Project Cedar', ['Maple by Maple', 'Project Cedar', 'NewProject Cedar'], 'Maple', true, 'https://example.org/demo']);
});
test('Delete and insert rows and columns project merged owners, text, formatting and continuations onto the surviving grid', async () => {
  const result = await run(() => {
    const results = [];
    for (const [op, r, c] of [['deleteCols', 2, 0], ['deleteCols', 2, 1], ['deleteRows', 0, 2], ['deleteRows', 1, 2], ['colLeft', 2, 1], ['rowAbove', 1, 2]]) {
      const p = L.model.newPresentation(), sl = p.slides[0], sh = L.model.newTable(p, sl, 3, 3, 20, 20, 540);
      sl.shapes = [sh];
      const vertical = /row/i.test(op), cell = sh.tbl.rows[0].cells[0];
      cell.tx.ps = [L.txt.para('Heading', {}, { b: true })]; cell.tx.vert = 'vert270';
      if (vertical) { cell.rs = 3; sh.tbl.rows[1].cells[0].vm = sh.tbl.rows[2].cells[0].vm = true; }
      else { cell.gs = 3; sh.tbl.rows[0].cells[1].hm = sh.tbl.rows[0].cells[2].hm = true; }
      L.app.loadPres(p, 'Table'); L.ed.select([sh.id]); L.ed.cellSel = { r0: r, c0: c, r1: r, c1: c }; L.app.tableCmd(op);
      const cells = vertical ? sh.tbl.rows.map(row => row.cells[0]) : sh.tbl.rows[0].cells;
      results.push([L.txt.plain(cells[0].tx), cells[0].tx.vert, vertical ? cells[0].rs : cells[0].gs, cells.slice(1).every(c => vertical ? c.vm : c.hm)]);
    }
    return results;
  });
  assert.deepEqual(result, [2, 2, 2, 2, 4, 4].map(n => ['Heading', 'vert270', n, true]));
});
test('Order preserves selected stacking order and moves each adjacent group by one unselected shape', async () => {
  const result = await run(() => {
    const out = [];
    for (const op of ['front', 'back', 'forward', 'backward']) {
      const p = L.model.newPresentation(), s = p.slides[0];
      s.shapes = Array.from('ABCDEF', name => ({ ...L.model.newShape(p, s, 'rect', 2, 2, 50, 50), name }));
      L.app.loadPres(p, 'Order'); L.ed.select([s.shapes[4], s.shapes[1], s.shapes[2]].map(sh => sh.id)); L.app.order(op);
      out.push(s.shapes.map(sh => sh.name).join(''));
    }
    return out;
  });
  assert.deepEqual(result, ['ADFBCE', 'BCEADF', 'ADBCFE', 'BCAEDF']);
});
test('Speaker note edits, undo and draft save retain links and rIds, paragraph levels, soft breaks, fields and page furniture', async () => {
  const result = await run(async () => {
    const p = L.model.newPresentation(); p.slides[0].notes = 'seed';
    const zip = await L.zip.read(new Uint8Array(await (await L.pptx.write(p)).arrayBuffer()));
    const entries = [], note = 'ppt/notesSlides/notesSlide1.xml', rel = L.opc.relsPath(note);
    for (const item of new Set(zip.values())) {
      let data = await item.bytes();
      if (item.name === note) data = (await item.text()).replace(/<a:p><a:r><a:rPr[^>]*\/><a:t>seed<\/a:t><\/a:r><\/a:p>/,
        '<a:p><a:pPr lvl="2"/><a:r><a:rPr b="1"><a:hlinkClick r:id="rId77"/></a:rPr><a:t>Demo</a:t></a:r><a:br/><a:fld id="{401DEDA2-F6D2-4848-9008-C1D63A4EF129}" type="datetime1"><a:t>Friday</a:t></a:fld></a:p>').replace('Slide Image Placeholder 1', 'Original page artwork');
      if (item.name === rel) data = (await item.text()).replace('</Relationships>', '<Relationship Id="rId77" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/demo" TargetMode="External"/></Relationships>');
      entries.push({ name: item.name, data });
    }
    const before = await L.pptx.read(new Uint8Array(await (await L.zip.write(entries)).arrayBuffer()));
    L.app.loadPres(before, 'Demo', { saved: true }); L.hist.clear();
    const box = document.querySelector('.notes-in'), original = box.textContent;
    box.textContent = original + ' reminder'; box.dispatchEvent(new Event('input'));
    L.hist.doUndo(); const undo = L.pres.slides[0].notes; L.hist.doRedo();
    const bytes = new Uint8Array(await (await __corpusHooks.snapshot(L.pres)).arrayBuffer()), pkg = await L.opc.open(bytes);
    const node = pkg.xml(note), by = name => [...node.getElementsByTagNameNS('*', name)];
    const again = await L.pptx.read(bytes);
    return [original, undo, again.slides[0].notes, by('hlinkClick')[0]?.getAttribute('r:id'), pkg.rels(note).find(r => r.id === 'rId77')?.target,
      by('pPr').some(x => x.getAttribute('lvl') === '2'), by('br').length, by('fld').some(x => x.getAttribute('id') === '{401DEDA2-F6D2-4848-9008-C1D63A4EF129}'),
      by('cNvPr').some(x => x.getAttribute('name') === 'Original page artwork')];
  });
  assert.deepEqual(result, ['Demo\u2028Friday', 'Demo\u2028Friday', 'Demo\u2028Friday reminder', 'rId77', 'https://example.org/demo', true, 1, true, true]);
});
