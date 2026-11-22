// A footnote that doesn't fit below its reference goes on to the next page, against Word 2016's prints of the same
// documents (tmp/lead-q/fn, 2026-11-07): Word places the reference's line, fills the rest of the page with as much of
// the note as fits (between paragraphs or inside one) and continues it at the top of the next page's notes, under a
// full-width line; the text after the reference waits for the next page. A note whose first paragraph keeps its lines
// together, or one with no room for a line, takes its reference's line to the next page instead (Quire did that for
// every note that didn't fit whole, and let a note longer than a page run off it).
// node --test tests/quire/footnotes-continued.test.mjs
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
const STYLES = '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>';
const para = (text) => '<w:p><w:r><w:t xml:space="preserve">' + text + '</w:t></w:r></w:p>';
const paras = (n, word) => Array.from({ length: n }, (_, i) => para(word + ' ' + i)).join('');
const REF = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
const refPara = (before, after = '') => '<w:p><w:r><w:t xml:space="preserve">' + before + '</w:t></w:r>' + REF + '<w:r><w:t xml:space="preserve">' + after + ' end.</w:t></w:r></w:p>';
const MARK = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>';
// a note of n one-line paragraphs, or of one long paragraph
const noteLines = (n) => Array.from({ length: n }, (_, i) => '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>' + (i ? '' : MARK) + '<w:r><w:t xml:space="preserve"> Note line ' + i + '</w:t></w:r></w:p>').join('');
const LONG = Array.from({ length: 180 }, (_, i) => 'word' + i).join(' ');
const notePara = (pPr = '') => '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/>' + pPr + '</w:pPr>' + MARK + '<w:r><w:t xml:space="preserve"> Start ' + LONG + ' Finish</w:t></w:r></w:p>';

// each page: its first and last line of text and the note text on it
async function pages(body, note) {
  const r = await page.evaluate(`(async () => {
    const N = L.opc.NS, part = 'word/document.xml', T = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
    const files = [
      { name: '[Content_Types].xml', data: '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/' + part + '" ContentType="' + T + 'document.main+xml"/><Override PartName="/word/styles.xml" ContentType="' + T + 'styles+xml"/><Override PartName="/word/footnotes.xml" ContentType="' + T + 'footnotes+xml"/></Types>' },
      { name: '_rels/.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/officeDocument" Target="' + part + '"/></Relationships>' },
      { name: 'word/_rels/document.xml.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/styles" Target="styles.xml"/><Relationship Id="rId2" Type="' + N.rel + '/footnotes" Target="footnotes.xml"/></Relationships>' },
      { name: 'word/styles.xml', data: '<w:styles xmlns:w="' + N.w + '">' + ${JSON.stringify(STYLES)} + '</w:styles>' },
      { name: 'word/footnotes.xml', data: '<w:footnotes xmlns:w="' + N.w + '"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">' + ${JSON.stringify(note)} + '</w:footnote></w:footnotes>' },
      { name: part, data: '<w:document xmlns:w="' + N.w + '"><w:body>' + ${JSON.stringify(body)} + '<w:sectPr>${SECT}</w:sectPr></w:body></w:document>' },
    ];
    await L.app.openFile(new File([await L.zip.write(files)], 'notes.docx'));
    await new Promise((r) => setTimeout(r, 300));
    return L.layout.pages.map((pg) => {
      const text = [...pg.el.querySelectorAll('.p[data-pid]')].filter((e) => !e.closest('.pg-fn')).map((e) => e.textContent.trim()).filter(Boolean);
      const notes = pg.el.querySelector('.pg-fn') ? pg.el.querySelector('.pg-fn').textContent.replace(/\\u200b/g, '').trim() : '';
      const box = pg.el.querySelector('.pg-fn'), body = pg.el.getBoundingClientRect();
      return { first: text[0], last: text[text.length - 1], notes, inside: !box || box.getBoundingClientRect().top >= body.top, cont: !!pg.el.querySelector('.fnsep-line.cont, .fnsep.cont') };
    });
  })()`);
  assert.deepEqual(page.errors, []);
  return r;
}

test('a note that does not fit below its reference goes on to the next page, between its paragraphs', async () => {
  // Word 2016: page 1 ends with the reference and note lines 0 to 9, page 2 starts with After 0 and lines 10 to 14
  const r = await pages(paras(36, 'Body') + refPara('Ref one') + paras(30, 'After'), noteLines(15));
  assert.equal(r.length, 2);
  assert.deepEqual([r[0].last, r[1].first], ['Ref one1 end.', 'After 0']);
  assert.match(r[0].notes, /line 0 .*line 9$/s);
  assert.match(r[1].notes, /^Note line 10 .*line 14$/s);
  assert.deepEqual([r[0].cont, r[1].cont], [false, true]);
});

test('a note goes on to the next page inside a paragraph, unless the paragraph keeps its lines together', async () => {
  // Word 2016: words 0 to 133 on page 1, the rest on page 2; with keep lines together, the reference's line moves
  const r = await pages(paras(36, 'Body') + refPara('Ref one') + paras(30, 'After'), notePara());
  assert.deepEqual([r.length, r[0].last, r[1].first], [2, 'Ref one1 end.', 'After 0']);
  assert.match(r[0].notes, /Start word0 /);
  assert.match(r[1].notes, /Finish$/);
  assert.doesNotMatch(r[1].notes, /Start/);
  const kept = await pages(paras(36, 'Body') + refPara('Ref one') + paras(30, 'After'), notePara('<w:keepLines/>'));
  assert.deepEqual([kept[0].last, kept[1].first, kept[0].notes], ['Body 35', 'Ref one1 end.', '']);
});

test('the text after the reference waits for the next page, but the reference does not stand alone at its foot', async () => {
  // Word 2016: the reference's paragraph keeps two lines on page 1 (widow and orphan control), the rest on page 2
  const long = Array.from({ length: 60 }, (_, i) => 'body' + i).join(' ');
  const r = await pages(paras(36, 'Body') + refPara('Ref line ', ' ' + long + ' Tail') + paras(30, 'After'), noteLines(15));
  assert.match(r[0].last, /^Ref line 1 body0 .* body23$/);
  assert.match(r[1].first, /^body24 /);
});

test('a note with no room for a line takes its reference to the next page; one longer than a page stays on its pages', async () => {
  // Word 2016: the reference on the last line moves with its note
  const r = await pages(paras(45, 'Body') + refPara('Ref one') + paras(30, 'After'), noteLines(4));
  assert.deepEqual([r[0].last, r[1].first, r[0].notes], ['Body 44', 'Ref one1 end.', '']);
  // Word 2016: three pages, the note's lines 0 to 46 on page 1, 47 to 69 on page 2 with After 0 to 25
  const huge = await pages(paras(5, 'Body') + refPara('Ref one') + paras(60, 'After'), noteLines(70));
  assert.equal(huge.length, 3);
  assert.ok(huge.every((p) => p.inside), 'the notes stay on the page');
  assert.match(huge[0].notes, /line 46$/);
  assert.match(huge[1].notes, /^Note line 47 .*line 69$/s);
  assert.deepEqual([huge[1].first, huge[1].last], ['After 0', 'After 25']);
});

// a one-column table, each row a paragraph, one row's paragraph referencing the note
const table = (...rows) => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  rows.map((r) => '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' + r + '</w:tc></w:tr>').join('') + '</w:tbl>';

test('a footnote referenced in a table cell is on the page, and goes on to the next one as from a paragraph', async () => {
  // Word 2016: the note under the table on page 1 (Quire left notes referenced in cells out altogether)
  const r = await pages(paras(30, 'Body') + table(para('Row 0'), refPara('Row 1 ref'), para('Row 2')) + paras(10, 'After'), noteLines(5));
  assert.match(r[0].notes, /line 0 .*line 4$/s);
  // Word 2016: Row 2 and note lines 0 to 2 on page 1, Row 3 and lines 3 and 4 on page 2
  const tight = await pages(paras(40, 'Body') + table(para('Row 0'), para('Row 1'), refPara('Row 2 ref'), para('Row 3')) + paras(10, 'After'), noteLines(5));
  assert.match(tight[0].notes, /line 0 .*line 2$/s);
  assert.deepEqual([tight[0].last, tight[1].first], ['Row 2 ref1 end.', 'Row 3']);
  assert.match(tight[1].notes, /^Note line 3 .*line 4$/s);
});
