// Paragraphs, runs and table rows keep their editing-session ids (w:rsid*) and paragraph ids (w14:paraId, w14:textId):
// Word 2016 places text per rsid run, so a save without them moved the text by fractions of a pixel (the Fluent
// Ribbon walkthrough), and Compare and Combine use them. A copied paragraph doesn't repeat its paragraph id. Text is
// written with xml:space="preserve" only where it needs it, as Word does.
// node --test tests/quire/run-ids.test.mjs
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
  const K = L.opc, N = K.NS, W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
  const body = '<w:p w:rsidR="00A1B2C3" w:rsidRDefault="00B2C3D4" w:rsidP="00C3D4E5" w14:paraId="1A2B3C4D" w14:textId="77777777">' +
    '<w:r w:rsidR="00D4E5F6" w:rsidRPr="00E5F6A7"><w:t xml:space="preserve">Hello </w:t></w:r><w:r w:rsidR="00F6A7B8"><w:t>world</w:t><w:tab/><w:t>x</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr w:rsidR="00A7B8C9" w:rsidTr="00B8C9D0" w14:paraId="2B3C4D5E" w14:textId="3C4D5E6F">' +
    '<w:tc><w:p w14:paraId="4D5E6F70" w14:textId="77777777"><w:r><w:t>Cell</w:t></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
  const comments = `<w:comments xmlns:w="${N.w}" xmlns:w14="${W14}"><w:comment w:id="0" w:author="A" w:initials="A">` +
    '<w:p w:rsidR="00C9D0E1" w14:paraId="0A0B0C0D" w14:textId="77777777"><w:r><w:t>Note</w:t></w:r></w:p></w:comment></w:comments>';
  const type = part => `application/vnd.openxmlformats-officedocument.wordprocessingml.${part}+xml`;
  const rel = (id, part, target) => `<Relationship Id="${id}" Type="${N.rel}/${part}" Target="${target}"/>`;
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="${type('document.main')}"/><Override PartName="/word/comments.xml" ContentType="${type('comments')}"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>` },
    { name: 'word/_rels/document.xml.rels', data: `<Relationships xmlns="${N.pkg}">${rel('rId1', 'comments', 'comments.xml')}</Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>` },
    { name: 'word/comments.xml', data: comments },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  if (kind === 'copied') doc.main.blocks.push(L.clone(doc.main.blocks[0]));
  // a copy pasted before its source, through the editor's copy (D.cloneBlocks)
  if (kind === 'copied-before') doc.main.blocks.unshift(...L.D.cloneBlocks([doc.main.blocks[0], doc.main.blocks[1]]));
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  const ids = el => Array.from(el.attributes).filter(a => /^(rsid|paraId|textId)/.test(a.localName)).map(a => a.localName + '=' + a.value).sort().join(' ');
  const roots = {}, all = (part, tag) => Array.from((roots[part] ||= saved.xml(part)).getElementsByTagNameNS(N.w, tag));
  return {
    paragraphs: all('word/document.xml', 'p').map(ids),
    runs: all('word/document.xml', 'r').filter(r => r.parentNode.localName === 'p' && r.parentNode === all('word/document.xml', 'p')[0]).map(ids),
    rows: all('word/document.xml', 'tr').map(ids),
    comment: all('word/comments.xml', 'p').map(ids),
    texts: all('word/document.xml', 't').map(t => [t.textContent, t.getAttribute('xml:space')]),
  };
}
const run = async kind => {
  const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(kind)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};
const first = 'paraId=1A2B3C4D rsidP=00C3D4E5 rsidR=00A1B2C3 rsidRDefault=00B2C3D4 textId=77777777';

test('an unedited save keeps the ids of paragraphs, runs and rows', async () => {
  const got = await run('unedited');
  assert.deepEqual(got.paragraphs, [first, 'paraId=4D5E6F70 textId=77777777', '']);
  // a run of several pieces is written as one run per piece, each with the run's ids
  assert.deepEqual(got.runs, ['rsidR=00D4E5F6 rsidRPr=00E5F6A7', 'rsidR=00F6A7B8', 'rsidR=00F6A7B8', 'rsidR=00F6A7B8']);
  assert.deepEqual(got.rows, ['paraId=2B3C4D5E rsidR=00A7B8C9 rsidTr=00B8C9D0 textId=3C4D5E6F']);
  assert.deepEqual(got.comment, ['paraId=0A0B0C0D rsidR=00C9D0E1 textId=77777777'], 'the comment\'s paragraph id once');
  // xml:space="preserve" only on text that needs it, as Word writes it (Word 2016 places text with it a fraction differently)
  assert.deepEqual(got.texts, [['Hello ', 'preserve'], ['world', null], ['x', null], ['Cell', null]]);
});

test('a copied paragraph keeps the editing-session ids but not the paragraph id', async () => {
  const got = await run('copied');
  assert.deepEqual(got.paragraphs, [first, 'paraId=4D5E6F70 textId=77777777', '', 'rsidP=00C3D4E5 rsidR=00A1B2C3 rsidRDefault=00B2C3D4']);
});

test('a copy pasted before its source leaves the paragraph and row ids with the source', async () => {
  const got = await run('copied-before');
  // the copy (Hello world, its table) first, then the source
  assert.deepEqual(got.paragraphs.slice(0, 2), ['rsidP=00C3D4E5 rsidR=00A1B2C3 rsidRDefault=00B2C3D4', '']);
  assert.deepEqual(got.paragraphs.slice(2, 4), [first, 'paraId=4D5E6F70 textId=77777777']);
  assert.deepEqual(got.rows, ['rsidR=00A7B8C9 rsidTr=00B8C9D0', 'paraId=2B3C4D5E rsidR=00A7B8C9 rsidTr=00B8C9D0 textId=3C4D5E6F']);
});
