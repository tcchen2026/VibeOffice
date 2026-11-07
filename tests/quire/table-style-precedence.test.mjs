// What a table style gives the paragraphs in its cells, against Word 2016's prints of documents made for the purpose
// (the Office VM, 2026-11-07): the table style comes after the document defaults and before the paragraph style,
// the Normal style included. What Normal sets itself (its size, colour, space after, justification, shading) wins over
// the table style, in compatibility modes 12, 14 and 15 and with overrideTableStyleFontSizeAndJustification on or off;
// what Normal leaves to the defaults the table style sets; a toggle Normal turns off (b=0) leaves the table style's
// bold on; below compatibility mode 15 Normal's own "left" gives way to the table style's justification ([MS-DOCX]
// overrideTableStyleFontSizeAndJustification). Quire let the table style override all of Normal's own settings, so
// LibreOffice's tdf90906_colAutoB.docx lost its red paragraph shading and 12 pt text in the cells.
// node --test tests/quire/table-style-precedence.test.mjs
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

/** each cell's first run: size (pt), colour, bold, its paragraph's alignment and background, and the gap (pt) from
    its first line to its second paragraph's */
async function cells(normal, tableStyle, compat = '') {
  const r = await page.evaluate(`(async () => {
    const K = L.opc, N = K.NS, part = 'word/document.xml', T = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
    const borders = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>';
    const styles = '<w:styles xmlns:w="' + N.w + '"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults>'
      + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' + ${JSON.stringify(normal)} + '</w:style>'
      + '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>'
      + '<w:style w:type="table" w:styleId="T1"><w:name w:val="T1"/><w:basedOn w:val="TableNormal"/>' + ${JSON.stringify(tableStyle)} + '<w:tblPr>' + borders + '</w:tblPr></w:style></w:styles>';
    const body = '<w:tbl><w:tblPr><w:tblStyle w:val="T1"/><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>'
      + '<w:p><w:r><w:t>Cell</w:t></w:r></w:p><w:p><w:r><w:t>Next</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
    const files = [
      { name: '[Content_Types].xml', data: '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/' + part + '" ContentType="' + T + 'document.main+xml"/><Override PartName="/word/styles.xml" ContentType="' + T + 'styles+xml"/><Override PartName="/word/settings.xml" ContentType="' + T + 'settings+xml"/></Types>' },
      { name: '_rels/.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/officeDocument" Target="' + part + '"/></Relationships>' },
      { name: 'word/_rels/document.xml.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/styles" Target="styles.xml"/><Relationship Id="rId2" Type="' + N.rel + '/settings" Target="settings.xml"/></Relationships>' },
      { name: 'word/styles.xml', data: styles },
      { name: 'word/settings.xml', data: '<w:settings xmlns:w="' + N.w + '"><w:compat>' + ${JSON.stringify(compat)} + '</w:compat></w:settings>' },
      { name: part, data: '<w:document xmlns:w="' + N.w + '"><w:body>' + body + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>' },
    ];
    await L.app.openFile(new File([await L.zip.write(files)], 'ts.docx'));
    L.layout.setZoom(1);
    await new Promise((r) => setTimeout(r, 300));
    const span = (t) => [...document.querySelectorAll('#pages td .pc > span')].find((s) => s.textContent === t);
    const a = span('Cell'), b = span('Next'), cs = getComputedStyle(a), pc = a.closest('.pc'), p = a.closest('.p');
    return { size: Math.round(parseFloat(cs.fontSize) * 0.75 * 10) / 10, color: cs.color, bold: +cs.fontWeight >= 600, align: getComputedStyle(pc).textAlign.replace('start', 'left'),
      shading: [getComputedStyle(p).backgroundColor, getComputedStyle(pc).backgroundColor].find((c) => c !== 'rgba(0, 0, 0, 0)') || 'none',
      gap: Math.round((b.getBoundingClientRect().top - a.getBoundingClientRect().top) * 0.75 * 10) / 10 };
  })()`);
  assert.deepEqual(page.errors, []);
  return r;
}
const MODE = (m) => '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="' + m + '"/>';
const OVERRIDE = '<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>';

test('what Normal sets itself wins over the table style, in every mode, the override on or off', async () => {
  const normal = '<w:pPr><w:jc w:val="right"/><w:shd w:val="clear" w:fill="FF0000"/><w:spacing w:after="400"/></w:pPr><w:rPr><w:sz w:val="24"/><w:color w:val="FF0000"/></w:rPr>';
  const table = '<w:pPr><w:jc w:val="center"/><w:shd w:val="clear" w:fill="auto"/><w:spacing w:after="0"/></w:pPr><w:rPr><w:sz w:val="32"/><w:color w:val="0000FF"/></w:rPr>';
  for (const compat of ['', MODE(14), MODE(15), MODE(15) + OVERRIDE]) {
    const r = await cells(normal, table, compat);
    // Word 2016: 12 pt red text, right-aligned on red, 20 pt after its line (13.8 pt) before the next paragraph
    assert.deepEqual([r.size, r.color, r.align, r.shading], [12, 'rgb(255, 0, 0)', 'right', 'rgb(255, 0, 0)'], compat || 'no mode');
    assert.ok(Math.abs(r.gap - 33.8) < 1.5, 'the space after Normal sets: ' + r.gap);
  }
});

test('what Normal leaves to the defaults the table style sets; Normal\'s b=0 leaves its bold on', async () => {
  const r = await cells('<w:rPr><w:b w:val="0"/></w:rPr>', '<w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="32"/><w:b/></w:rPr>', MODE(15));
  assert.deepEqual([r.size, r.align, r.bold], [16, 'center', true]);
});

test('below compatibility mode 15 Normal\'s own "left" gives way to the table style\'s justification', async () => {
  // [MS-DOCX] overrideTableStyleFontSizeAndJustification; Word 2016: right-aligned with no mode and in mode 14,
  // left in mode 15 and with the setting on; Normal's "right" holds in any mode (tdf118947_tableStyle.docx)
  const normal = '<w:pPr><w:jc w:val="left"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr>', table = '<w:pPr><w:jc w:val="right"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr>';
  assert.equal((await cells(normal, table, '')).align, 'right', 'no mode');
  assert.equal((await cells(normal, table, MODE(14))).align, 'right', 'mode 14');
  assert.equal((await cells(normal, table, MODE(15))).align, 'left', 'mode 15');
  assert.equal((await cells(normal, table, MODE(14) + OVERRIDE)).align, 'left', 'mode 14, the setting on');
  assert.equal((await cells('<w:pPr><w:jc w:val="center"/></w:pPr>', table, '')).align, 'center', 'Normal centred');
  assert.equal((await cells(normal, table, '')).size, 12, 'Normal\'s 12 pt over 11 pt defaults holds');
});
