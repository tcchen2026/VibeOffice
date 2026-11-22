// File ▸ Properties ▸ Custom: a property added in the dialog ({ type, v }) is saved with its value and its type, as
// Word writes them (vt:lpwstr, vt:i4 or vt:r8, vt:bool, vt:filetime), and one read from a file (a string) as text.
// It used to be saved as "[object Object]".
// node --test tests/quire/custom-properties.test.mjs
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
  const K = L.opc, N = K.NS;
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body><w:p><w:r><w:t>Text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>` },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  doc.custom = { Client: 'Acme', Pages: { type: 'number', v: 12 }, Ratio: { type: 'number', v: 1.5 }, Approved: { type: 'bool', v: true }, Due: { type: 'date', v: '2026-11-04' }, Owner: { type: 'text', v: 'Lead' } };
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  const x = saved.xml('docProps/custom.xml');
  return [...x.getElementsByTagNameNS(x.documentElement.namespaceURI, 'property')].map((p) => [p.getAttribute('name'), p.firstElementChild.localName, p.firstElementChild.textContent]);
}

test('custom properties keep their values and types', async () => {
  const got = await page.evaluate(`(${save.toString()})()`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  assert.deepEqual(got, [
    ['Client', 'lpwstr', 'Acme'], ['Pages', 'i4', '12'], ['Ratio', 'r8', '1.5'], ['Approved', 'bool', 'true'],
    ['Due', 'filetime', '2026-11-04T00:00:00Z'], ['Owner', 'lpwstr', 'Lead'],
  ]);
});
