import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

let page, stop;
before(async () => { stop = await ensureServer(); page = await openPage('quire', { persistence: false, automated: true }); });
after(async () => { try { await page?.close(); } finally { stop?.(); } });
async function run(fn) {
  const result = await page.evaluate('(' + fn.toString() + ')()');
  assert.deepEqual(page.errors, []);
  return result;
}

test('pasting notes and comment threads clones their bodies and IDs in the same undo step, including Insert File', async () => {
  const rows = await run(async () => {
    const { D, ed: E } = L, rows = [];
    function documentWith(name) {
      const d = D.newDoc(), id = '2', item = (kind, id = '2') => D.item(kind, { id });
      for (const kind of ['fn', 'en']) d[kind][id] = { id, kind, blocks: [D.para([D.item(kind, { self: true, id }), D.text(name + kind, { b: true })])] };
      d.comments[2] = { id, kind: 'cmt', author: name, blocks: [D.para([D.text(name + ' comment')])] };
      d.comments[3] = { id: '3', kind: 'cmt', author: name, parent: id, blocks: [D.para([D.text(name + ' reply')])] };
      d.main.blocks = [D.para([item('cs'), item('cs', '3'), D.text(name), item('fn'), item('en'), item('ce', '3'), item('ce')])];
      return d;
    }
    for (const mode of ['same', 'cross', 'file']) {
      const src = documentWith('Source'), dst = mode === 'same' ? src : documentWith('Other');
      L.app.loadDoc(src, 'Source');
      const srcP = src.main.blocks[0]; E.setSel(D.pos(srcP, 0), D.pos(srcP, D.plen(srcP)));
      const payload = E.copyData();
      L.app.loadDoc(dst, 'Destination');
      const p = dst.main.blocks[0]; E.setSel(D.pos(p, D.plen(p)));
      if (mode === 'file') {
        const file = new File([await L.docx.write(src)], 'source.docx'), pick = L.pickFiles;
        L.pickFiles = async () => [file];
        try { await L.app.insertFileDialog(); } finally { L.pickFiles = pick; }
      } else E.pasteBlocks(payload.blocks, { bodies: payload.bodies });
      const again = (await L.docx.read(new Uint8Array(await (await L.docx.write(dst)).arrayBuffer()))).doc;
      const refs = kind => again.main.blocks.flatMap(b => b.runs.filter(it => it.t === kind).map(it => it.id));
      const fn = refs('fn'), en = refs('en'), threads = Object.values(again.comments);
      const last = (store, ids) => store[ids.at(-1)]?.blocks[0].runs.find(it => it.t === 'text');
      const undo = D.hist.undo.length;
      D.undo();
      rows.push([new Set(fn).size, new Set(en).size, last(again.fn, fn)?.text, last(again.en, en)?.rPr.b, threads.length,
        threads.filter(c => c.parent != null).map(c => again.comments[c.parent]?.author === c.author), undo,
        Object.keys(dst.fn).length, Object.keys(dst.en).length, Object.keys(dst.comments).length]);
    }
    return rows;
  });
  assert.deepEqual(rows, Array.from({ length: 3 }, () => [2, 2, 'Sourcefn', true, 4, [true, true], 1, 1, 1, 2]));
});

test('older clipboard payloads drop orphan markers with a notice and inserted footnotes skip retained separator IDs', async () => {
  const result = await run(() => {
    const { D, ed: E } = L, d = D.newDoc(), p = d.main.blocks[0], msg = [];
    d.keep = { notes: { fn: [{ id: '0' }, { id: '1' }] } };
    d.fn['2'] = { id: '2', kind: 'fn', blocks: [D.para([D.text('Different citation')])] };
    L.app.loadDoc(d, 'Report');
    const toast = L.ui.toast;
    try {
      L.ui.toast = text => msg.push(text);
      E.pasteBlocks([D.para([D.text('Citation'), D.item('fn', { id: '2' })])], { bodies: undefined });
      L.notes.insert('fn');
    } finally { L.ui.toast = toast; }
    return [p.runs.filter(it => it.t === 'fn').map(it => it.id), msg.length, Object.keys(d.fn)];
  });
  assert.deepEqual(result, [['3'], 1, ['2', '3']]);
});

test('Find, Highlight and Replace use revision visibility, keep hidden offsets and never edit a deletion', async () => {
  const result = await run(() => {
    const { D, find, ed: E, layout } = L, d = D.newDoc();
    const p = D.para([D.text('Mon'), D.text('day', { del: { author: 'A' } }), D.text(' week'), D.text('new', { ins: { author: 'A' } })]);
    d.main.blocks = [p]; L.app.loadDoc(d, 'Schedule');
    const hits = text => find.all({ text }, d.main).matches.map(m => [m.a.o, m.b.o]);
    layout.opts.markup = 'final'; layout.render();
    const final = [hits('Mon'), hits('day'), hits(' week'), find.highlightAll({ text: 'day' })];
    E.setSel(D.pos(p, 0)); const next = find.next({ text: ' week' }).a.o;
    layout.opts.markup = 'original'; const original = [hits('day'), hits('new')];
    layout.opts.markup = 'markup'; E.setSel(D.pos(p, 3), D.pos(p, 6));
    find.replaceOne({ text: 'day', replace: 'Tuesday' });
    const skipped = find.replaceAll({ text: 'day', replace: 'Tuesday' });
    layout.opts.markup = 'final';
    const changed = find.replaceAll({ text: 'Mon week', replace: 'Meeting' });
    layout.opts.markup = 'markup'; find.clearHighlight();
    return [final, next, original, skipped, changed, p.runs.map(it => [it.text, !!it.rPr.del])];
  });
  assert.deepEqual(result, [[[[0, 3]], [], [[6, 11]], 0], 6, [[[3, 6]], []], 0, 1, [['Meeting', false], ['day', true], ['new', false]]]);
});

test('Print Selection lays out just the selected table text and note, preserves page numbers and restores the editor', async () => {
  const result = await run(async () => {
    const { D, ed: E, layout: ly } = L, d = D.newDoc(), rows = [], table = D.simpleTable(2, 3, 480);
    const ps = table.rows.flatMap(r => r.cells.map(c => c.blocks[0]));
    ps.forEach((p, i) => { p.runs = [D.text('private' + i)]; });
    ps[1].runs = [D.text('prefixPUBLIC'), D.item('fn', { id: '2' }), D.text('deleted', { del: { author: 'A' } })];
    ps[2].runs = [D.item('fb', { fid: 'page', instr: ' PAGE ' }), D.item('fs', { fid: 'page' }), D.text('1'), D.item('fe', { fid: 'page' }), D.text('ENDsuffix')];
    d.fn[2] = { kind: 'fn', id: '2', blocks: [D.para([D.text('Cited source')])] };
    d.sect.pgNum = { start: 7 }; d.main.blocks = [D.para([D.text('Private pricing')]), table, D.para([D.text('Confidential')])];
    L.app.loadDoc(d, 'Excerpt'); ly.opts.markup = 'final'; ly.render();
    E.setSel(D.pos(ps[1], 6), D.pos(ps[2], 4));
    const build = L.app.buildPDF, save = L.saveFile, dialog = L.ui.printDialog;
    try {
      for (const mode of ['doc', 'markup']) {
        let done;
        const saved = new Promise(resolve => { done = resolve; });
        L.saveFile = async () => done();
        L.app.buildPDF = async () => {
          const pg = ly.pages[0], text = pg.el.textContent;
          rows.push([D.doc.main.blocks.map(b => b.t), pg.numVal, /private|prefix|suffix|Confidential/i.test(text), text.includes('PUBLIC'),
            text.includes('Cited source'), text.includes('deleted'), text.includes('END')]);
          return new Blob();
        };
        L.ui.printDialog = opts => { opts.what.querySelector('#pr-what').value = mode; opts.print({ range: 'selection' }); };
        L.dlg.print(); await saved;
        rows.at(-1).push(D.doc === d, E.sel.a.p === ps[1] && E.sel.f.p === ps[2], ly.opts.markup);
      }
    } finally { L.app.buildPDF = build; L.saveFile = save; L.ui.printDialog = dialog; }
    return rows;
  });
  assert.deepEqual(result, [false, true].map(markup => [['tbl'], 7, false, true, true, markup, true, true, true, 'final']));
});

test('merge labels and preview use successive included records, conditional rules, record numbers and empty trailing labels', async () => {
  const result = await run(() => {
    const { D, mailmerge: mm } = L, d = D.newDoc(); let number = 0;
    const field = instr => { const fid = String(++number); return [D.item('fb', { instr, fid }), D.item('fs', { fid }), D.text('cached'), D.item('fe', { fid })]; };
    const table = D.simpleTable(1, 2, 400);
    const left = table.rows[0].cells[0].blocks[0], right = table.rows[0].cells[1].blocks[0];
    left.runs = [...field('SKIPIF { MERGEFIELD Name } = "Skip"'), ...field('MERGEREC'), D.text(' '), ...field('MERGEFIELD Name')];
    right.runs = [...field('NEXTIF { MERGEFIELD Advance } = "yes"'), ...field('MERGESEQ'), D.text(' '), ...field('MERGEFIELD Name')];
    d.main.blocks = [table]; L.app.loadDoc(d, 'Labels'); mm.preview = false; mm.rec = 0; mm.type = 'labels';
    mm.data = { fields: ['Name', 'Advance'], rows: [['Ada','yes'],['Excluded','yes'],['Ben','yes'],['Skip','yes'],['Cia','yes'],['Dee','yes'],['Eve','no'],['Flo','yes']], included: new Set([0,2,3,4,5,6,7]) };
    mm.go(0); const preview = [D.plainText(left), D.plainText(right)];
    const read = doc => doc.main.blocks.filter(b => b.t === 'tbl').map(t => t.rows[0].cells.map(c => D.plainText(c.blocks[0])));
    const conditional = read(mm.buildMerged());
    right.runs[0].instr = 'NEXT'; const next = read(mm.buildMerged());
    mm.data = null; mm.preview = false;
    return [preview, conditional, next];
  });
  assert.deepEqual(result, [['1 Ada','2 Ben'], [['1 Ada','2 Ben'],['4 Cia','5 Dee'],['6 Eve','6 Eve'],['7 Flo',' ']],
    [['1 Ada','2 Ben'],['4 Cia','5 Dee'],['6 Eve','7 Flo']]]);
});

test('Word table formulas read percentage operands as fractions and numeric pictures print a literal percent sign', async () => {
  const result = await run(() => {
    const { D, tables: T } = L, d = D.newDoc(), t = D.simpleTable(2, 3, 440);
    const p = t.rows[1].cells[2].blocks[0];
    for (const [r, c, text] of [[0,0,'400%'],[1,0,'100'],[1,1,'20.00%']]) t.rows[r].cells[c].blocks[0].runs = [D.text(text)];
    d.main.blocks = [t]; L.app.loadDoc(d, 'Invoice');
    return [T.evalFormula('A2*B2', p), T.evalFormula('AVERAGE(A1)', p), T.parseNum('20.00%'),
      L.fields.evaluate('=50*40 \\# "0%"', D.pos(p,0)), T.formatNumber(-2, '0%;(0%)')];
  });
  assert.deepEqual(result, [20, 4, 20, '2000%', '(2%)']);
});

test('a copied comment that crosses the selection boundary has two anchors and one new Word reference', async () => {
  const result = await run(async () => {
    const D = L.D, d = D.newDoc(), E = L.ed;
    const p = D.para([D.text('A'), D.item('cs', { id: '0' }), D.text('BCD'), D.item('ce', { id: '0' }), D.text('E')]);
    d.comments[0] = { kind: 'cmt', id: '0', author: 'Editor', blocks: [D.para([D.text('Review')])] };
    d.main.blocks = [p]; L.app.loadDoc(d, 'Comments');
    for (const [a, b] of [[0, 2], [3, 5]]) {
      E.setSel(D.pos(p, a), D.pos(p, b));
      const data = E.copyData(); E.setSel(D.pos(p, D.plen(p)));
      E.pasteBlocks(data.blocks, { bodies: data.bodies });
    }
    const pkg = await L.opc.open(new Uint8Array(await (await L.docx.write(d)).arrayBuffer()));
    const xml = pkg.xml('word/document.xml');
    const ids = tag => Array.from(xml.getElementsByTagNameNS('*', tag), el => el.getAttribute('w:id')).sort();
    return [ids('commentRangeStart'), ids('commentRangeEnd'), ids('commentReference')];
  });
  assert.deepEqual(result, Array.from({ length: 3 }, () => ['0', '1', '2']));
});

test('a Word NEXTIF with a nested MERGEFIELD evaluates its current record rather than its cached result', async () => {
  const result = await run(async () => {
    const D = L.D, pkg = await L.zip.read(new Uint8Array(await (await L.docx.write(D.newDoc())).arrayBuffer()));
    const run = x => '<w:r>' + x + '</w:r>';
    const ch = type => run('<w:fldChar w:fldCharType="' + type + '"/>');
    const code = s => run('<w:instrText xml:space="preserve">' + s + '</w:instrText>');
    const field = (instr, cached) => ch('begin') + instr + ch('separate') + run('<w:t>' + cached + '</w:t>') + ch('end');
    const name = field(code(' MERGEFIELD Name '), 'cached');
    const conditional = field(code(' NEXTIF "') + field(code(' MERGEFIELD Advance '), 'wrong') + code('" = "yes" '), '');
    const entries = [];
    for (const part of new Set(pkg.values())) {
      let data = await part.bytes();
      if (part.name === 'word/document.xml') data = (await part.text()).replace(/<w:body>[\s\S]*?<w:sectPr/, '<w:body><w:p>' + name + conditional + run('<w:t>/</w:t>') + name + '</w:p><w:sectPr');
      entries.push({ name: part.name, data });
    }
    const doc = (await L.docx.read(new Uint8Array(await (await L.zip.write(entries)).arrayBuffer()))).doc;
    L.app.loadDoc(doc, 'Nested rule');
    const mm = L.mailmerge; mm.preview = false; mm.type = 'labels';
    mm.data = { fields: ['Name','Advance'], rows: [['Ada','yes'],['Ben','no'],['Cia','no']], included: new Set([0,1,2]) };
    mm.go(0);
    const preview = D.plainText(doc.main.blocks[0]), merged = mm.buildMerged().main.blocks.map(p => D.plainText(p));
    mm.data = null; mm.preview = false;
    return [preview, merged];
  });
  assert.deepEqual(result, ['Ada/Ben', ['Ada/Ben', 'Cia/Cia']]);
});

test('opening a damaged older paste gives each mark an editable note and a Word-safe save', async () => {
  const result = await run(async () => {
    const D = L.D, d = D.newDoc(), p = d.main.blocks[0], id = '2';
    for (const kind of ['fn','en']) {
      d[kind][id] = { kind, id, blocks: [D.para([D.item(kind, { id, self: true }), D.text(kind + ' body')])] };
      p.runs.push(D.text(' marker'), D.item(kind, { id }), D.item(kind, { id }));
    }
    const broken = await L.zip.read(new Uint8Array(await (await L.docx.write(d)).arrayBuffer())), entries = [];
    for (const part of new Set(broken.values())) {
      let data = await part.bytes();
      if (part.name === 'word/document.xml') data = (await part.text()).replace('</w:p>', '<w:r><w:footnoteReference w:id="99"/></w:r></w:p>');
      entries.push({ name: part.name, data });
    }
    const { doc, warnings } = await L.docx.read(new Uint8Array(await (await L.zip.write(entries)).arrayBuffer()));
    const saved = await L.opc.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
    const refs = kind => Array.from(saved.xml('word/document.xml').getElementsByTagNameNS('*', kind + 'Reference'), el => el.getAttribute('w:id'));
    return [refs('footnote'), refs('endnote'), Object.values(doc.fn).map(st => st.blocks[0].runs.at(-1).text), warnings.some(w => /footnote/i.test(w))];
  });
  assert.deepEqual(result, [['2','3'], ['2','3'], ['fn body','fn body'], true]);
});
