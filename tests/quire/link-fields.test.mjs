// What an unedited save used to drop without a word: a simple field's lock, a link field's ScreenTip with quotes, a link's target frame.
// node --test tests/quire/link-fields.test.mjs
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

async function roundTrip() {
  const K = L.opc, N = K.NS, part = 'word/document.xml';
  const field = type => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
  const body = '<w:p><w:fldSimple w:instr=" PRINTDATE \\* MERGEFORMAT " w:fldLock="1"><w:r><w:t>1/23/4567</w:t></w:r></w:fldSimple></w:p>' +
    `<w:p>${field('begin')}<w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com/" \\o "First \\"link\\" here" </w:instrText></w:r>${field('separate')}<w:r><w:t>First</w:t></w:r>${field('end')}</w:p>` +
    `<w:p>${field('begin')}<w:r><w:instrText xml:space="preserve">HYPERLINK "https://example.com/" \\o "Unclosed \\"tip\\" </w:instrText></w:r>${field('separate')}<w:r><w:t>Second</w:t></w:r>${field('end')}</w:p>` +
    '<w:p><w:hyperlink r:id="rIdLink" w:tgtFrame="_blank"><w:r><w:t>Elsewhere</w:t></w:r></w:hyperlink></w:p>';
  const relationships = content => `<Relationships xmlns="${N.pkg}">${content}</Relationships>`;
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: '_rels/.rels', data: relationships(`<Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="${part}"/>`) },
    { name: K.relsPath(part), data: relationships(`<Relationship Id="rIdLink" Type="${N.rel}/hyperlink" Target="https://example.org/" TargetMode="External"/>`) },
    { name: part, data: `<w:document xmlns:w="${N.w}" xmlns:r="${N.rel}"><w:body>${body}<w:sectPr/></w:body></w:document>` },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  const root = saved.xml(part), all = tag => Array.from(root.getElementsByTagNameNS(N.w, tag));
  return {
    begins: all('fldChar').filter(e => e.getAttributeNS(N.w, 'fldCharType') === 'begin').map(e => e.getAttributeNS(N.w, 'fldLock')),
    links: all('hyperlink').map(e => [e.getAttributeNS(N.w, 'tooltip'), e.getAttributeNS(N.w, 'tgtFrame')]),
  };
}

let result;
const run = async () => {
  result ||= await page.evaluate(`(${roundTrip.toString()})()`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

test('a locked simple field stays locked', async () => {
  assert.deepEqual((await run()).begins, ['1'], 'PRINTDATE keeps its lock');
});

test('links keep a ScreenTip with quotes and the frame they open in', async () => {
  // a missing closing quote: the tip runs to the end of the code, as LibreOffice reads it
  assert.deepEqual((await run()).links, [['First "link" here', null], ['Unclosed "tip"', null], [null, '_blank']]);
});
