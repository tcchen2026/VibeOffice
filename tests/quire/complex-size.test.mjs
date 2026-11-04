// A run read with a size (w:sz) but no complex-script size (w:szCs) is saved without one: writing szCs = sz there
// changed the size of Arabic, Hebrew and other complex-script text (27,025 times in 683 Word-authored files). Text
// Quire formats itself still gets both, as Word writes them.
// node --test tests/quire/complex-size.test.mjs
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

async function save() {
  const K = L.opc, N = K.NS, D = L.D;
  const body = '<w:p><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>Read</w:t></w:r><w:r><w:rPr><w:sz w:val="28"/><w:szCs w:val="32"/></w:rPr><w:t xml:space="preserve"> both</w:t></w:r></w:p>';
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body>${body}<w:sectPr/></w:body></w:document>` },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  doc.main.blocks.push(D.para([D.text('New', { sz: 9 })]));
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  return Array.from(saved.xml('word/document.xml').getElementsByTagNameNS(N.w, 'r')).map(r => {
    const v = tag => r.getElementsByTagNameNS(N.w, tag)[0]?.getAttributeNS(N.w, 'val') ?? null;
    return [r.getElementsByTagNameNS(N.w, 't')[0]?.textContent, v('sz'), v('szCs')];
  });
}

test('a run keeps the sizes it was read with; new text gets both', async () => {
  const got = await page.evaluate(`(${save.toString()})()`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  assert.deepEqual(got, [['Read', '28', null], [' both', '28', '32'], ['New', '18', '18']]);
});
