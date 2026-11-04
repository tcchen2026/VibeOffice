// The comment extensions Quire adds on save (word/commentsExtended.xml: threads, resolved state) use the w15 namespace the
// file's parts bind: Word 2016 refused files from Word 2013's preview, which bind the draft one, when they got the released
// one (Comment012, Comment035, Comment050 of the Open XML SDK's tests; tools/ooxml/package.py, w15-namespaces).
// node --test tests/quire/comment-extensions.test.mjs
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

async function save(w15) {
  const K = L.opc, N = K.NS;
  const comment = `<w:comment w:id="0" w:author="A" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment>`;
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/_rels/document.xml.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/comments" Target="comments.xml"/></Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}" xmlns:w15="${w15}"><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr/></w:body></w:document>` },
    { name: 'word/comments.xml', data: `<w:comments xmlns:w="${N.w}" xmlns:w15="${w15}">${comment}</w:comments>` },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  return saved.xml('word/commentsExtended.xml')?.documentElement.namespaceURI ?? null;
}
const run = async w15 => {
  const got = await page.evaluate(`(${save.toString()})(${JSON.stringify(w15)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return got;
};

test('a file binding the draft w15 namespace gets its comment extensions in the draft one', async () => {
  const draft = 'http://schemas.microsoft.com/office/word/2010/11/wordml';
  assert.equal(await run(draft), draft);
});

test('any other file gets them in the released one', async () => {
  const released = 'http://schemas.microsoft.com/office/word/2012/wordml';
  assert.equal(await run(released), released);
});
