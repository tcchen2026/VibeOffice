// Styles a file has under ids of its own (Word in German writes Überschrift1 for heading 1) are not written a second time.
// node --test tests/quire/style-names.test.mjs
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
  const K = L.opc, N = K.NS, D = L.D;
  const style = (type, id, name, body = '', def = '') => `<w:style w:type="${type}"${def} w:styleId="${id}"><w:name w:val="${name}"/>${body}</w:style>`;
  const styles = `<w:styles xmlns:w="${N.w}">` + style('paragraph', 'Standard', 'Normal', '<w:rPr><w:sz w:val="24"/></w:rPr>', ' w:default="1"') +
    style('character', 'Absatz-Standardschriftart', 'Default Paragraph Font', '', ' w:default="1"') +
    style('paragraph', 'berschrift1', 'heading 1', '<w:basedOn w:val="Standard"/><w:next w:val="Standard"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr>') +
    style('paragraph', 'Textkrper', 'Body Text', '<w:basedOn w:val="Standard"/><w:rPr><w:b/></w:rPr>') +
    style('paragraph', 'Textkrper2', 'Body Text 2', '<w:basedOn w:val="Standard"/><w:pPr><w:jc w:val="both"/></w:pPr>') +
    style('character', 'Seitenzahl', 'page number', '<w:basedOn w:val="Absatz-Standardschriftart"/>') +
    style('table', 'Tabellenraster', 'Table Grid', '<w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>') + '</w:styles>';
  const para = (id, text) => `<w:p><w:pPr><w:pStyle w:val="${id}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const body = para('berschrift1', 'Die Grundsätze') + para('Textkrper2', 'Ziel der Initiative') + '<w:p><w:r><w:t>Text</w:t></w:r></w:p>';
  const type = kind => `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`;
  const rel = (id, kind, target) => `<Relationship Id="${id}" Type="${N.rel}/${kind}" Target="${target}"/>`;
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="${type('document.main')}"/><Override PartName="/word/styles.xml" ContentType="${type('styles')}"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>` },
    { name: 'word/_rels/document.xml.rels', data: `<Relationships xmlns="${N.pkg}">${rel('rId1', 'styles', 'styles.xml')}</Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>` },
    { name: 'word/styles.xml', data: styles },
  ]);
  const read = async () => (await L.docx.read(new Uint8Array(await original.arrayBuffer()))).doc;
  const written = async doc => {
    const pkg = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
    const root = pkg.xml('word/styles.xml');
    written.paragraphs = Array.from(pkg.xml('word/document.xml').getElementsByTagNameNS(N.w, 'pStyle')).map(e => e.getAttributeNS(N.w, 'val'));
    written.tables = Array.from(pkg.xml('word/document.xml').getElementsByTagNameNS(N.w, 'tblStyle')).map(e => e.getAttributeNS(N.w, 'val'));
    return Array.from(root.getElementsByTagNameNS(N.w, 'style')).map(s => ({ id: s.getAttributeNS(N.w, 'styleId'), type: s.getAttributeNS(N.w, 'type'),
      name: s.getElementsByTagNameNS(N.w, 'name')[0].getAttributeNS(N.w, 'val').toLowerCase(),
      basedOn: s.getElementsByTagNameNS(N.w, 'basedOn')[0]?.getAttributeNS(N.w, 'val') || null }));
  };
  const unedited = await written(await read());
  if (!D.useStyle) return { unedited };
  const doc = await read();
  const used = { heading: D.useStyle(doc, 'Heading1'), pageNumber: D.useStyle(doc, 'PageNumber'), toc: D.useStyle(doc, 'TOC1') };
  // a command that sets a built-in by id without asking (the paragraph's look in Quire is the built-in's until reopened)
  const paragraphs = []; D.walk(doc.main, b => { if (b.t === 'p') paragraphs.push(b); });
  paragraphs[2].pPr.style = 'BodyText2';
  // built-ins the file doesn't define, put to use
  for (const id of [used.toc, 'BodyTextFirstIndent', 'TOCHeading']) doc.main.blocks.push(D.para([D.text(id)], { style: id }));
  // a new table is given Table Grid by id, as Insert > Table does
  doc.main.blocks.push(D.simpleTable(2, 2, 200));
  const edited = await written(doc);
  return { unedited, used, edited, paragraphs: written.paragraphs, tables: written.tables };
}
let result;
const run = async () => {
  result ||= await page.evaluate(`(${roundTrip.toString()})()`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

const duplicates = list => {
  const seen = new Map();
  for (const s of list) seen.set(s.type + ':' + s.name, (seen.get(s.type + ':' + s.name) || []).concat(s.id));
  return [...seen.values()].filter(ids => ids.length > 1);
};

test('an unedited save writes each style name once, under the file\'s own ids', async () => {
  const { unedited } = await run();
  assert.deepEqual(duplicates(unedited), [], 'two styles of one name');
  const ids = unedited.map(s => s.id);
  for (const id of ['berschrift1', 'Textkrper', 'Textkrper2', 'Seitenzahl']) assert.ok(ids.includes(id), id);
  for (const id of ['Heading1', 'BodyText', 'BodyText2', 'PageNumber']) assert.ok(!ids.includes(id), id + ' is the file\'s ' + id.replace(/\d$/, ' $&'));
  // built-ins the file doesn't define and nothing uses are left to its latent styles, as Word does
  for (const id of ['BodyTextFirstIndent', 'TOCHeading', 'TOC1']) assert.ok(!ids.includes(id), id + ' is not added');
});

test('applying a built-in style uses the file\'s style of that name, and adds one the file doesn\'t have', async () => {
  const result = await run();
  assert.deepEqual(result.used, { heading: 'berschrift1', pageNumber: 'Seitenzahl', toc: 'TOC1' });
  assert.deepEqual(duplicates(result.edited), []);
  assert.ok(result.edited.some(s => s.id === 'TOC1' && s.name === 'toc 1'));
  // a paragraph given the built-in by id is written with the file's style of that name, as Word applies its own
  assert.deepEqual(result.paragraphs, ['berschrift1', 'Textkrper2', 'Textkrper2', 'TOC1', 'BodyTextFirstIndent', 'TOCHeading']);
  // a built-in put to use inherits from the file's style of the name it was based on
  assert.equal(result.edited.find(s => s.id === 'BodyTextFirstIndent')?.basedOn, 'Textkrper');
  assert.equal(result.edited.find(s => s.id === 'TOCHeading')?.basedOn, 'berschrift1');
  assert.deepEqual(result.tables, ['Tabellenraster']);
  assert.ok(result.edited.some(s => s.id === 'Tabellenraster') && !result.edited.some(s => s.id === 'TableGrid'));
});
