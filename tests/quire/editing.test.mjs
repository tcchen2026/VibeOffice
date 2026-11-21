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
