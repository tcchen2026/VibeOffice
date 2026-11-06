// Where paragraphs stand, against Word 2016's prints of the same documents (tmp/qp/sect and tmp/qp/auto, 2026-11-06):
// - an empty paragraph that only holds a continuous section break takes no room at all, neither a line nor its spacing,
//   and the space after and before on either side of the break collapse to the larger, as between any paragraphs
//   (LibreOffice's tdf169986_bottomSpacing.docx, one page in Word, took two in Quire);
// - auto spacing stays out from between paragraphs whose numbering goes on from one to the next (tdf133052.docx,
//   one page in Word, took two; bullet-simple1.docx keeps it between two lists).
// node --test tests/quire/paragraph-spacing.test.mjs
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

const SECT = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>';
const para = (text, pPr = '') => '<w:p><w:pPr>' + pPr + '</w:pPr>' + (text ? '<w:r><w:t>' + text + '</w:t></w:r>' : '') + '</w:p>';
const breakPara = (text, pPr) => para(text, pPr + '<w:sectPr>' + SECT + '</w:sectPr>');
const AUTO = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>';
const item = (text, numId, ilvl = 0) => para(text, '<w:numPr><w:ilvl w:val="' + ilvl + '"/><w:numId w:val="' + numId + '"/></w:numPr>' + AUTO);
const LEVEL = (i, fmt) => '<w:lvl w:ilvl="' + i + '"><w:start w:val="1"/><w:numFmt w:val="' + fmt + '"/><w:lvlText w:val="' + (fmt === 'bullet' ? 'o' : '%' + (i + 1) + '.') + '"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="' + 720 * (i + 1) + '" w:hanging="360"/></w:pPr></w:lvl>';
const NUMBERING = '<w:abstractNum w:abstractNumId="0">' + LEVEL(0, 'decimal') + LEVEL(1, 'lowerLetter') + '</w:abstractNum><w:abstractNum w:abstractNumId="1">' + LEVEL(0, 'bullet') + '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="0"/></w:num>' +
  '<w:num w:numId="4"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>';

// opens the body (Times New Roman 12 pt, single spacing, the last section continuous) and returns how far below
// Alpha's line (pt) each paragraph's text stands, by its first word
async function offsets(body) {
  const r = await page.evaluate(`(async () => {
    const K = L.opc, N = K.NS, part = 'word/document.xml', T = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
    const files = [
      { name: '[Content_Types].xml', data: '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/' + part + '" ContentType="' + T + 'document.main+xml"/><Override PartName="/word/styles.xml" ContentType="' + T + 'styles+xml"/><Override PartName="/word/numbering.xml" ContentType="' + T + 'numbering+xml"/></Types>' },
      { name: '_rels/.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/officeDocument" Target="' + part + '"/></Relationships>' },
      { name: 'word/_rels/document.xml.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/styles" Target="styles.xml"/><Relationship Id="rId2" Type="' + N.rel + '/numbering" Target="numbering.xml"/></Relationships>' },
      { name: 'word/styles.xml', data: '<w:styles xmlns:w="' + N.w + '"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>' },
      { name: 'word/numbering.xml', data: '<w:numbering xmlns:w="' + N.w + '">' + ${JSON.stringify(NUMBERING)} + '</w:numbering>' },
      { name: part, data: '<w:document xmlns:w="' + N.w + '"><w:body>' + ${JSON.stringify(body)} + '<w:sectPr><w:type w:val="continuous"/>${SECT}</w:sectPr></w:body></w:document>' },
    ];
    await L.app.openFile(new File([await L.zip.write(files)], 'spacing.docx'));
    L.layout.setZoom(1);
    await new Promise((r) => setTimeout(r, 300));
    const tops = {};
    for (const s of document.querySelectorAll('#pages .pc > span:not(.lbl)')) { const t = s.textContent.trim(); if (t && !(t in tops)) tops[t] = s.getBoundingClientRect().top; }
    const out = {};
    for (const [t, y] of Object.entries(tops)) out[t] = (y - tops.Alpha) * 0.75;
    return out;
  })()`);
  assert.deepEqual(page.errors, []);
  return r;
}

const near = (got, want, what) => assert.ok(Math.abs(got - want) <= 1.5, `${what}: ${got.toFixed(1)} pt below Alpha, Word ${want}`);

test('an empty paragraph holding a continuous section break takes no room, neither its line nor its spacing', async () => {
  // Word 2016: Beta 33.9 pt below Alpha (one line and the 20 pt both sides ask for); the break's 100 pt after is not there
  let r = await offsets(para('Alpha', '<w:spacing w:after="400"/>') + breakPara('', '<w:spacing w:after="2000"/>') + para('Beta', '<w:spacing w:before="400"/>') + para('Gamma'));
  near(r.Beta, 33.9, 'Beta after a break with 100 pt after');
  // nor a line as tall as its 48 pt mark, nor its space before: Beta one line below Alpha
  r = await offsets(para('Alpha') + breakPara('', '<w:spacing w:before="2000"/><w:rPr><w:sz w:val="96"/></w:rPr>') + para('Beta') + para('Gamma'));
  near(r.Beta, 13.8, 'Beta after a break with a 48 pt mark and 100 pt before');
});

test('space after and space before collapse across a continuous section break', async () => {
  // Word 2016: Mid's 100 pt after and Beta's 20 pt before leave 100 pt between them, not 120
  const r = await offsets(para('Alpha', '<w:spacing w:after="400"/>') + breakPara('Mid', '<w:spacing w:after="2000"/>') + para('Beta', '<w:spacing w:before="400"/>') + para('Gamma'));
  near(r.Mid, 34.0, 'Mid');
  near(r.Beta, 147.6, 'Beta');
});

test('auto spacing stays out from between paragraphs whose numbering goes on, at any level, but not next to others', async () => {
  // Word 2016, one line 13.8 pt: 14 pt before the first item; none to the next level, nor to another numbering of the
  // same list that goes on counting (numId 3); 14 pt to another list (bullets), to a restart and to a plain paragraph
  const r = await offsets(para('Alpha') + item('One', 1) + item('Two', 1, 1) + item('Three', 3) + item('Four', 2) + item('Five', 4) + para('Plain', AUTO) + para('Omega'));
  near(r.One, 27.8, 'the first item');
  near(r.Two, 41.6, 'an item of the next level');
  near(r.Three, 55.4, 'an item of a numbering that goes on counting');
  near(r.Four, 83.2, 'an item of another list');
  near(r.Five, 111.0, 'an item that restarts the numbering');
  near(r.Plain, 138.8, 'a paragraph after the list');
  near(r.Omega, 166.6, 'the paragraph after that');
});

test('paragraphs with the same borders and no indent share one box, as tall as Word makes it', async () => {
  // Word 2016 (tdf135343): a box's top border, its two lines (Liberation Sans, as tall as Arial: 12.6 pt for 11 pt) with
  // the 6 pt between them, and its bottom border; the next paragraph 39.1 pt below the first line
  const BOX = '<w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr><w:spacing w:before="120" w:after="120"/>';
  const font = '<w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/><w:sz w:val="22"/></w:rPr>';
  const run = (t) => '<w:r>' + font + '<w:t>' + t + '</w:t></w:r>';
  const r = await offsets('<w:p><w:pPr>' + BOX + '</w:pPr>' + run('Alpha') + '</w:p><w:p><w:pPr>' + BOX + '</w:pPr>' + run('Boxed') + '</w:p><w:p>' + run('After') + '</w:p>');
  near(r.Boxed, 18.6, 'the second line of the box');
  near(r.After, 39.1, 'the paragraph after the box');
});
