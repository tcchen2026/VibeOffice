// A damaged body with two final section definitions: Word 2016 and LibreOffice lay out the last section with the last
// one (tdf108849: page 2 landscape), so Quire shows and saves that one.
// node --test tests/quire/final-section.test.mjs
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
  const K = L.opc, N = K.NS;
  // tdf108849's shape: a landscape first section, then two final sections between the paragraphs, the last one landscape
  const body = '<w:p><w:pPr><w:sectPr><w:pgSz w:w="8504" w:h="5669" w:orient="landscape"/></w:sectPr></w:pPr><w:r><w:t>Page 1</w:t></w:r></w:p>' +
    '<w:sectPr><w:pgSz w:w="5669" w:h="8504"/></w:sectPr><w:sectPr><w:pgSz w:w="14173" w:h="5669" w:orient="landscape"/></w:sectPr>' +
    '<w:p><w:r><w:t>Page 2</w:t></w:r></w:p>';
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body>${body}</w:body></w:document>` },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  const finals = Array.from(saved.xml('word/document.xml').getElementsByTagNameNS(N.w, 'sectPr')).filter(e => e.parentNode.localName === 'body');
  const size = e => { const s = e.getElementsByTagNameNS(N.w, 'pgSz')[0]; return [s.getAttributeNS(N.w, 'w'), s.getAttributeNS(N.w, 'h'), s.getAttributeNS(N.w, 'orient') || null]; };
  return { shown: [doc.sect.pgW, doc.sect.pgH, doc.sect.orient || null], saved: finals.map(size) };
}

test('the last of two final section definitions is the one shown and saved', async () => {
  const got = await page.evaluate(`(${roundTrip.toString()})()`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  assert.deepEqual(got.shown, [708.65, 283.45, 'landscape']);
  assert.deepEqual(got.saved, [['14173', '5669', 'landscape']]);
});
