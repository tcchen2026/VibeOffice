// A legacy form field's settings survive a save: help and status text with their line breaks, entry and exit
// macros, Calculate on exit and a disabled field (I24 E2: converted Word 97-2003 forms lost them).
// node --test tests/quire/form-fields.test.mjs
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
  const ff = '<w:ffData><w:name w:val="Gender"/><w:enabled w:val="0"/><w:calcOnExit w:val="1"/><w:entryMacro w:val="text_entry"/><w:exitMacro w:val="text_exit"/>' +
    '<w:helpText w:type="text" w:val="Gender of the candidate&#13;&#10;e.g.  M or F&#13;&#10;"/><w:statusText w:type="text" w:val="One&#9;letter"/><w:textInput><w:maxLength w:val="1"/></w:textInput></w:ffData>';
  const runs = `<w:r><w:fldChar w:fldCharType="begin">${ff}</w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>M</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const files = [
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="${part}"/></Relationships>` },
    { name: part, data: `<w:document xmlns:w="${N.w}"><w:body><w:p>${runs}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>` },
  ];
  const { doc } = await L.docx.read(new Uint8Array(await (await L.zip.write(files)).arrayBuffer()));
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  // read as Word reads it: a conforming parser turns a line break written literally inside an attribute into a space
  const text = new TextDecoder().decode(saved.bytes(saved.main));
  const data = new DOMParser().parseFromString(text, 'application/xml').getElementsByTagNameNS(N.w, 'ffData')[0];
  return [...data.children].map((e) => [e.localName, e.getAttributeNS(N.w, 'val')]);
}

test('a form field keeps its help text\'s line breaks, its macros and its settings when saved', async () => {
  const kids = await page.evaluate(`(${roundTrip.toString()})()`);
  assert.deepEqual(page.errors, []);
  const val = (name) => (kids.find(([k]) => k === name) || [])[1];
  assert.equal(val('helpText'), 'Gender of the candidate\r\ne.g.  M or F\r\n');
  assert.equal(val('statusText'), 'One\tletter');
  assert.equal(val('entryMacro'), 'text_entry');
  assert.equal(val('exitMacro'), 'text_exit');
  assert.equal(val('calcOnExit'), '1');
  assert.equal(val('enabled'), '0');
  assert.deepEqual(kids.map(([k]) => k).slice(0, 7), ['name', 'enabled', 'calcOnExit', 'entryMacro', 'exitMacro', 'helpText', 'statusText'], 'in the order Word writes them');
});
