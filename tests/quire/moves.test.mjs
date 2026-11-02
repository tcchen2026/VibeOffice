// A tracked move (moveFrom, moveTo and their named ranges) is saved as a move, not as a deletion and an insertion.
// node --test tests/quire/moves.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

let page, stop;
before(async () => {
  stop = await ensureServer(process.env.BASE);
  page = await openPage('quire', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

async function save(kind) {
  const K = L.opc, N = K.NS, D = L.D;
  const by = 'w:author="A" w:date="2026-01-01T00:00:00Z"';
  // Word's shape: a paragraph moved down past another, its mark in the move, each range ending after its paragraph
  const body = `<w:p><w:pPr><w:rPr><w:moveFrom w:id="10" ${by}/></w:rPr></w:pPr><w:moveFromRangeStart w:id="1" ${by} w:name="move1"/>` +
    `<w:moveFrom w:id="2" ${by}><w:r><w:t>Moved</w:t></w:r></w:moveFrom></w:p><w:moveFromRangeEnd w:id="1"/>` +
    '<w:p><w:r><w:t>Stays</w:t></w:r></w:p>' +
    `<w:p><w:pPr><w:rPr><w:moveTo w:id="11" ${by}/></w:rPr></w:pPr><w:moveToRangeStart w:id="3" ${by} w:name="move1"/>` +
    `<w:moveTo w:id="4" ${by}><w:r><w:t>Moved</w:t></w:r></w:moveTo></w:p><w:moveToRangeEnd w:id="3"/>` +
    '<w:p><w:r><w:t>End</w:t></w:r></w:p>';
  const type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="${type}"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body>${body}<w:sectPr/></w:body></w:document>` },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  D.doc = doc; D.resetHistory();
  const paragraphs = []; D.walk(doc.main, b => { if (b.t === 'p') paragraphs.push(b); });
  if (kind === 'accepted') D.tx('Accept', () => {
    // the moved-to half accepted on its own, as Quire's review pane does for an insertion
    const p = paragraphs[2]; D.touch(p);
    for (const it of p.runs) if (it.rPr?.ins) delete it.rPr.ins;
    delete p.mark.ins;
  });
  if (kind === 'copied') D.tx('Copy', () => { doc.main.blocks.push(...L.preserve.duplicateControls([paragraphs[0]])); });
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  const root = saved.xml('word/document.xml'), sequence = [], ids = [];
  const text = el => ['t', 'delText'].flatMap(tag => Array.from(el.getElementsByTagNameNS(N.w, tag))).map(t => t.textContent).join('');
  const walk = (el, inMark) => {
    for (const c of el.children) {
      const tag = c.localName;
      if (/^(ins|del|move(From|To)(Range(Start|End))?)$/.test(tag)) {
        sequence.push([tag, inMark ? 'mark' : c.parentNode.localName, c.getAttributeNS(N.w, 'name') || text(c)].join(' ').trim());
        if (!/End$/.test(tag)) ids.push(c.getAttributeNS(N.w, 'id'));
      }
      walk(c, inMark || tag === 'rPr' && c.parentNode.localName === 'pPr');
    }
  };
  walk(root.getElementsByTagNameNS(N.w, 'body')[0]);
  const losses = (doc.losses || []).filter(e => e.id.startsWith('move:')).map(e => ({ id: e.id.replace(/:[^:]*:word/, ':word'), notify: !!e.notify }));
  return { sequence, unique: new Set(ids).size === ids.length, losses };
}
const run = async kind => {
  const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(kind)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

test('an unedited save keeps the move, its paragraph marks and named ranges where they were', async () => {
  const { sequence, unique, losses } = await run('unedited');
  assert.deepEqual(sequence, ['moveFrom mark', 'moveFromRangeStart p move1', 'moveFrom p Moved', 'moveFromRangeEnd body',
    'moveTo mark', 'moveToRangeStart p move1', 'moveTo p Moved', 'moveToRangeEnd body']);
  assert.ok(unique, 'revision and range ids are unique');
  assert.deepEqual(losses, []);
});

test('a move whose other half was accepted is saved as the deletion left, without its ranges', async () => {
  const { sequence, losses } = await run('accepted');
  assert.deepEqual(sequence, ['del mark', 'del p Moved']);
  assert.deepEqual(losses, [{ id: 'move:word/document.xml:moveFrom:1', notify: false }, { id: 'move:word/document.xml:moveTo:3', notify: false }]);
});

test('a copy of moved text is an ordinary deletion; the move itself stays', async () => {
  const { sequence, unique } = await run('copied');
  assert.deepEqual(sequence.slice(0, 8), ['moveFrom mark', 'moveFromRangeStart p move1', 'moveFrom p Moved', 'moveFromRangeEnd body',
    'moveTo mark', 'moveToRangeStart p move1', 'moveTo p Moved', 'moveToRangeEnd body']);
  assert.deepEqual(sequence.slice(8), ['del mark', 'del p Moved']);
  assert.ok(unique);
});
