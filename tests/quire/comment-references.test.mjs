// A comment's reference run keeps its properties and ids, and its place after the other comments' marks: Quire wrote
// every reference right after its range's end as a plain CommentReference run, so Word 2016 placed the text after it
// a fraction of a point differently (Comment035, Comment050; 505 of the corpus's 1,364 reference runs have ids, 325 more
// properties, 232 come after other comments' marks, some right before their range's end).
// node --test tests/quire/comment-references.test.mjs
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

async function save(edit) {
  const K = L.opc, N = K.NS, D = L.D;
  const ref = (id, pr = '', attrs = '') => `<w:r${attrs}><w:rPr><w:rStyle w:val="CommentReference"/>${pr}</w:rPr><w:commentReference w:id="${id}"/></w:r>`;
  // three comments ending together, as Word writes them: two ends, the first reference, the third end, the other two
  const body = '<w:p><w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/><w:commentRangeStart w:id="2"/><w:r><w:t>Text</w:t></w:r>' +
    `<w:commentRangeEnd w:id="0"/><w:commentRangeEnd w:id="1"/>${ref(0, '<w:sz w:val="16"/>', ' w:rsidRPr="00AB12CD"')}<w:commentRangeEnd w:id="2"/>${ref(1)}${ref(2)}` +
    '<w:r><w:t xml:space="preserve"> more</w:t></w:r></w:p>' +
    // a reference Word wrote right before its range's end
    `<w:p><w:commentRangeStart w:id="3"/><w:r><w:t>Second</w:t></w:r>${ref(3, '', ' w:rsidRPr="00AB12CE"')}<w:commentRangeEnd w:id="3"/></w:p>`;
  const comment = id => `<w:comment w:id="${id}" w:author="A" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Note ${id}</w:t></w:r></w:p></w:comment>`;
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/_rels/document.xml.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/comments" Target="comments.xml"/></Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body>${body}<w:sectPr/></w:body></w:document>` },
    { name: 'word/comments.xml', data: `<w:comments xmlns:w="${N.w}">${comment(0)}${comment(1)}${comment(2)}${comment(3)}</w:comments>` },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  const p = doc.main.blocks[0];
  if (edit === 'typed') p.runs.splice(p.runs.findIndex(it => it.t === 'ce') + 1, 0, D.text('typed'));
  if (edit === 'deleted') { delete doc.comments[1]; p.runs = p.runs.filter(it => it.id !== '1'); }
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  const w = (el, tag) => el.getElementsByTagNameNS(N.w, tag)[0];
  return Array.from(saved.xml('word/document.xml').getElementsByTagNameNS(N.w, 'p')).map(p => Array.from(p.children).map(el => {
    if (el.localName === 'commentRangeEnd') return 'end ' + el.getAttributeNS(N.w, 'id');
    if (el.localName !== 'r') return null;
    const r = w(el, 'commentReference');
    if (!r) return w(el, 't').textContent.trim();
    return ['ref ' + r.getAttributeNS(N.w, 'id'), el.getAttributeNS(N.w, 'rsidRPr'), w(el, 'sz')?.getAttributeNS(N.w, 'val')].filter(Boolean).join(' ');
  }).filter(Boolean));
}
const run = async edit => {
  const got = await page.evaluate(`(${save.toString()})(${JSON.stringify(edit)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return got;
};

test('an unedited save writes the references where Word had them, with their properties and ids', async () => {
  assert.deepEqual(await run(), [['Text', 'end 0', 'end 1', 'ref 0 00AB12CD 16', 'end 2', 'ref 1', 'ref 2', 'more'], ['Second', 'ref 3 00AB12CE', 'end 3']]);
});

test('text typed after a range\'s end gets its reference before it', async () => {
  assert.deepEqual((await run('typed'))[0], ['Text', 'end 0', 'ref 0 00AB12CD 16', 'typed', 'end 1', 'end 2', 'ref 1', 'ref 2', 'more']);
});

test('a deleted comment takes its marks; the others\' references still follow the ends', async () => {
  // comment 2 is saved as 1
  assert.deepEqual((await run('deleted'))[0], ['Text', 'end 0', 'end 1', 'ref 0 00AB12CD 16', 'ref 1', 'more']);
});
