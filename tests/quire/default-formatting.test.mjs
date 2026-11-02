// A file that names no default formatting gets Word's (Calibri 11 pt, 8 pt after, 1.08 lines), measured in Word 2016;
// one that names a default but leaves a value out keeps the format's own (Times New Roman, 10 pt, single).
// node --test tests/quire/default-formatting.test.mjs
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

async function defaults(styles) {
  const K = L.opc, N = K.NS;
  const type = kind => `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`;
  const rel = (id, kind, target) => `<Relationship Id="${id}" Type="${N.rel}/${kind}" Target="${target}"/>`;
  const files = [
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="${type('document.main')}"/>${styles ? `<Override PartName="/word/styles.xml" ContentType="${type('styles')}"/>` : ''}</Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body><w:p><w:r><w:t>Some text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>` },
  ];
  if (styles) files.push({ name: 'word/_rels/document.xml.rels', data: `<Relationships xmlns="${N.pkg}">${rel('rId1', 'styles', 'styles.xml')}</Relationships>` },
    { name: 'word/styles.xml', data: `<w:styles xmlns:w="${N.w}">${styles}<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>` });
  const { doc } = await L.docx.read(new Uint8Array(await (await L.zip.write(files)).arrayBuffer()));
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  const root = saved.xml('word/styles.xml'), first = (tag) => root.getElementsByTagNameNS(N.w, 'docDefaults')[0].getElementsByTagNameNS(N.w, tag)[0];
  return { font: first('rFonts')?.getAttributeNS(N.w, 'ascii'), size: first('sz')?.getAttributeNS(N.w, 'val'),
    after: first('spacing')?.getAttributeNS(N.w, 'after') || null, line: first('spacing')?.getAttributeNS(N.w, 'line') || null };
}
const run = async (styles) => {
  const result = await page.evaluate(`(${defaults.toString()})(${JSON.stringify(styles)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

test('no styles part, or no defaults in it: Word\'s Calibri 11 pt, 8 pt after, 1.08 lines', async () => {
  const word = { font: 'Calibri', size: '22', after: '160', line: '259' };
  assert.deepEqual(await run(null), word);
  assert.deepEqual(await run(''), word);
  assert.deepEqual(await run('<w:docDefaults/>'), word);
});

test('a default that is there keeps the format\'s values for what it leaves out', async () => {
  assert.deepEqual(await run('<w:docDefaults><w:rPrDefault/><w:pPrDefault/></w:docDefaults>'), { font: 'Times New Roman', size: '20', after: null, line: null });
  assert.deepEqual(await run('<w:docDefaults><w:pPrDefault/></w:docDefaults>'), { font: 'Calibri', size: '22', after: null, line: null });
  assert.deepEqual(await run('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:rPrDefault></w:docDefaults>'), { font: 'Arial', size: '20', after: '160', line: '259' });
});
