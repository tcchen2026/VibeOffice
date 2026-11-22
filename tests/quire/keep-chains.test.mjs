// Paragraphs kept with the next one ("keep with next") through a table, against Word 2016's prints of the same
// documents (tmp/lead-q/keep, 2026-11-06): a table whose last row keeps with the next paragraph (its first cell's first
// paragraph does, whatever the other cells and paragraphs say) stays in the chain of such paragraphs, so a chain that
// starts at the top of a page is laid out from there and not moved to the next page (POI's bug65649.docx, where every
// paragraph keeps with the next and tables break the chains: Word 177 pages, Quire 181); a row whose first cell's
// first paragraph doesn't keep ends the chain, and the paragraphs after the table move on together.
// node --test tests/quire/keep-chains.test.mjs
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
const para = (text, keep) => '<w:p><w:pPr>' + (keep ? '<w:keepNext/>' : '') + '<w:spacing w:after="0"/></w:pPr><w:r><w:t>' + text + '</w:t></w:r></w:p>';
const row = (...cells) => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>' + cells.map(() => '<w:gridCol w:w="3000"/>').join('') + '</w:tblGrid><w:tr>' +
  cells.map((c) => '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>' + c + '</w:tc>').join('') + '</w:tr></w:tbl>';
const keeps = (n) => Array.from({ length: n }, (_, i) => para('Keep ' + i, true)).join('');
const after5 = Array.from({ length: 5 }, (_, i) => para('After ' + i)).join('');
// a page, then a page break, a title and the table at the top of page 2, and 70 paragraphs that keep with the next
const doc = (table) => para('Plain') + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + para('Title', true) + table + keeps(70) + after5;

// the first paragraph on each page, by its text
async function pageStarts(body) {
  const r = await page.evaluate(`(async () => {
    const N = L.opc.NS, part = 'word/document.xml';
    const files = [
      { name: '[Content_Types].xml', data: '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/' + part + '" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>' },
      { name: '_rels/.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/officeDocument" Target="' + part + '"/></Relationships>' },
      { name: part, data: '<w:document xmlns:w="' + N.w + '"><w:body>' + ${JSON.stringify(body)} + '<w:sectPr>${SECT}</w:sectPr></w:body></w:document>' },
    ];
    await L.app.openFile(new File([await L.zip.write(files)], 'keep.docx'));
    await new Promise((r) => setTimeout(r, 300));
    return L.layout.pages.map((pg) => { const p = pg.el.querySelector('.p[data-pid]'); return p ? p.textContent.trim() : ''; });
  })()`);
  assert.deepEqual(page.errors, []);
  return r;
}

test('a table whose row keeps with the next paragraph stays in the chain, which is not moved from the top of a page', async () => {
  // Word 2016: page 2 holds the title, the table and Keep 0 to 41
  assert.deepEqual(await pageStarts(doc(row(para('Cell', true)))), ['Plain', 'Title', 'Keep 42']);
  // only the first cell's first paragraph counts: Word 2016 the same for both
  assert.equal((await pageStarts(doc(row(para('A', true), para('B')))))[2], 'Keep 42');
  assert.equal((await pageStarts(doc(row(para('A1', true) + para('A2'), para('B', true)))))[2], 'Keep 41');
});

test('a table whose row does not keep with the next paragraph ends the chain, and the chain after it moves on', async () => {
  // Word 2016: page 2 holds only the title and the table; Keep 0 starts page 3
  assert.deepEqual(await pageStarts(doc(row(para('Cell')))), ['Plain', 'Title', 'Keep 0', 'Keep 44']);
  assert.deepEqual(await pageStarts(doc(row(para('A'), para('B', true)))), ['Plain', 'Title', 'Keep 0', 'Keep 44']);
});
