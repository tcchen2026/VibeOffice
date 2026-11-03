// The styles part keeps what Quire doesn't model: the latent styles, each style's unhideWhenUsed, locked and rsid, a table
// style's row properties. A built-in style is written once something uses it or it is changed, with its latent settings,
// as Word does; the others aren't added. Found on the Fluent Ribbon walkthrough (Heading 2–6 lost unhideWhenUsed).
// node --test tests/quire/styles-part.test.mjs
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
  const K = L.opc, N = K.NS;
  const latent = '<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="0" w:defUnhideWhenUsed="0" w:defQFormat="0" w:count="371">' +
    '<w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/><w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>' +
    '<w:lsdException w:name="heading 2" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>' +
    '<w:lsdException w:name="heading 3" w:semiHidden="1" w:uiPriority="9" w:unhideWhenUsed="1" w:qFormat="1"/>' +
    '<w:lsdException w:name="Title" w:uiPriority="10" w:qFormat="1"/></w:latentStyles>';
  const style = (type, id, name, body, extra = '') => `<w:style w:type="${type}"${extra} w:styleId="${id}"><w:name w:val="${name}"/>${body}</w:style>`;
  const styles = `<w:styles xmlns:w="${N.w}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>${latent}` +
    style('paragraph', 'Normal', 'Normal', '<w:qFormat/><w:rsid w:val="00A1B2C3"/>', ' w:default="1"') +
    style('paragraph', 'Heading1', 'heading 1', '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:rsid w:val="00A1B2C3"/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr>') +
    style('paragraph', 'Heading2', 'heading 2', '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:semiHidden/><w:unhideWhenUsed/><w:qFormat/><w:locked/><w:rsid w:val="00D4E5F6"/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:i/></w:rPr>') +
    style('character', 'DefaultParagraphFont', 'Default Paragraph Font', '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/>', ' w:default="1"') +
    style('table', 'TableNormal', 'Normal Table', '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/></w:tblPr>', ' w:default="1"') +
    style('table', 'Rows', 'Rows', '<w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:trPr><w:cantSplit/></w:trPr>' +
      '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tblPr/><w:trPr><w:tblHeader/></w:trPr></w:tblStylePr>', ' w:customStyle="1"') +
    style('numbering', 'NoList', 'No List', '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>', ' w:default="1"') +
    // a file already at Word 2016's limit of 4,092 styles
    (kind === 'full' ? Array.from({ length: 4092 - 7 }, (_, i) => style('paragraph', 'Pad' + i, 'Pad ' + i, '<w:basedOn w:val="Normal"/>', ' w:customStyle="1"')).join('') : '') + '</w:styles>';
  const para = (id, text) => `<w:p><w:pPr><w:pStyle w:val="${id}"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const body = para('Heading1', 'One') + para('Heading2', 'Two') + '<w:p><w:r><w:t>Text</w:t></w:r></w:p>' +
    '<w:tbl><w:tblPr><w:tblStyle w:val="Rows"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
  const type = part => `application/vnd.openxmlformats-officedocument.wordprocessingml.${part}+xml`;
  const rel = (id, part, target) => `<Relationship Id="${id}" Type="${N.rel}/${part}" Target="${target}"/>`;
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="${type('document.main')}"/><Override PartName="/word/styles.xml" ContentType="${type('styles')}"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}">${rel('rId1', 'officeDocument', 'word/document.xml')}</Relationships>` },
    { name: 'word/_rels/document.xml.rels', data: `<Relationships xmlns="${N.pkg}">${rel('rId1', 'styles', 'styles.xml')}</Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body>${body}<w:sectPr/></w:body></w:document>` },
    { name: 'word/styles.xml', data: styles },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  if (kind === 'full') doc.main.blocks[2].pPr = { ...doc.main.blocks[2].pPr, style: 'Heading3' };
  if (kind === 'edited') {
    // Heading 3, which the file doesn't define, applied to the plain paragraph; Heading 1 made italic
    doc.main.blocks[2].pPr = { ...doc.main.blocks[2].pPr, style: 'Heading3' };
    doc.styles.Heading1 = { ...doc.styles.Heading1, rPr: { ...doc.styles.Heading1.rPr, i: true } };
    // Title, which the file doesn't define either, changed (Modify Style) but used nowhere
    doc.styles.Title = { ...doc.styles.Title, rPr: { ...doc.styles.Title.rPr, color: 'C00000' } };
  }
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  const root = saved.xml('word/styles.xml'), w = (el, tag) => Array.from(el.getElementsByTagNameNS(N.w, tag));
  const children = id => {
    const s = w(root, 'style').find(e => e.getAttributeNS(N.w, 'styleId') === id);
    return s && Array.from(s.children).map(c => c.localName + (c.getAttributeNS(N.w, 'val') ? '=' + c.getAttributeNS(N.w, 'val') : '') + (['pPr', 'rPr', 'trPr', 'tblStylePr'].includes(c.localName) ? '(' + Array.from(c.children).map(g => g.localName).join(',') + ')' : ''));
  };
  return {
    ids: w(root, 'style').map(e => e.getAttributeNS(N.w, 'styleId')),
    latent: w(root, 'lsdException').length,
    Heading1: children('Heading1'), Heading2: children('Heading2'), Heading3: children('Heading3'), Rows: children('Rows'),
    limit: (doc.losses || []).filter(e => e.id.startsWith('styles-limit:')).map(e => ({ notify: !!e.notify, what: e.what })),
  };
}
const run = async kind => {
  const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(kind)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

test('an unedited save keeps the latent styles, what each style has that Quire doesn\'t model, and adds no style', async () => {
  const got = await run('unedited');
  assert.deepEqual(got.ids, ['Normal', 'Heading1', 'Heading2', 'DefaultParagraphFont', 'TableNormal', 'Rows', 'NoList'], 'the file\'s styles in its order, nothing added');
  assert.equal(got.latent, 5);
  assert.deepEqual(got.Heading2, ['name=heading 2', 'basedOn=Normal', 'next=Normal', 'uiPriority=9', 'semiHidden', 'unhideWhenUsed', 'qFormat', 'locked', 'rsid=00D4E5F6', 'pPr(keepNext,outlineLvl)', 'rPr(i)']);
  assert.deepEqual(got.Rows.slice(-2), ['trPr(cantSplit)', 'tblStylePr(rPr,tblPr,trPr)'], 'a table style\'s row properties');
});

test('an edited style keeps the rest; a built-in put to use comes with its latent settings, a changed one is written too', async () => {
  const got = await run('edited');
  assert.deepEqual(got.Heading1, ['name=heading 1', 'basedOn=Normal', 'next=Normal', 'uiPriority=9', 'qFormat', 'rsid=00A1B2C3', 'pPr(keepNext,outlineLvl)', 'rPr(b,i)']);
  assert.deepEqual(got.Heading3?.slice(0, 7), ['name=heading 3', 'basedOn=Normal', 'next=Normal', 'uiPriority=9', 'semiHidden', 'unhideWhenUsed', 'qFormat']);
  assert.deepEqual(got.ids.filter(id => !['Normal', 'Heading1', 'Heading2', 'DefaultParagraphFont', 'TableNormal', 'Rows', 'NoList'].includes(id)), ['Heading3', 'Title'], 'only the built-ins put to use or changed are added');
});

test('a file at Word\'s limit of 4,092 styles stays there: a built-in put to use is left out, and the user is told', async () => {
  const got = await run('full');
  assert.equal(got.ids.length, 4092);
  assert.equal(got.Heading3, undefined);
  assert.deepEqual(got.limit, [{ notify: true, what: 'Word opens at most 4,092 styles, so 1 style (heading 3) was left out; text using it shows in the Normal style in Word.' }]);
});
