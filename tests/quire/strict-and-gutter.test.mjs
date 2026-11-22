// Two page geometry readings, against Word 2016's prints (2026-11-07):
// - Strict files give "auto" line spacing as a measure (line="12.95pt" is 259 twips, 1.08 lines); Quire read 12.95
//   twips, so the lines of the SDK's Strict01.docx ran into each other (8 pages; Word 13, Quire 13 now);
// - a gutter at the top of the page (w:gutterAtTop) adds to the top margin, not the left (the SDK's Business Plan for a
//   Startup Business: Word's text starts half an inch lower; 29 pages, Word 31, Quire 30 now).
// node --test tests/quire/strict-and-gutter.test.mjs
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

// opens a document (body, sectPr's margins, settings) and returns probe's result
async function open(body, mar, settings, probe) {
  const r = await page.evaluate(`(async () => {
    const N = L.opc.NS, part = 'word/document.xml', T = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
    const files = [
      { name: '[Content_Types].xml', data: '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/' + part + '" ContentType="' + T + 'document.main+xml"/><Override PartName="/word/settings.xml" ContentType="' + T + 'settings+xml"/></Types>' },
      { name: '_rels/.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/officeDocument" Target="' + part + '"/></Relationships>' },
      { name: 'word/_rels/document.xml.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/settings" Target="settings.xml"/></Relationships>' },
      { name: 'word/settings.xml', data: '<w:settings xmlns:w="' + N.w + '">' + ${JSON.stringify(settings)} + '</w:settings>' },
      { name: part, data: '<w:document xmlns:w="' + N.w + '"><w:body>' + ${JSON.stringify(body)} + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar ' + ${JSON.stringify(mar)} + '/></w:sectPr></w:body></w:document>' },
    ];
    await L.app.openFile(new File([await L.zip.write(files)], 'geometry.docx'));
    L.layout.setZoom(1);
    await new Promise((r) => setTimeout(r, 300));
    return (${probe})();
  })()`);
  assert.deepEqual(page.errors, []);
  return r;
}
const MAR = 'w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="720"';
const para = (t, pPr = '') => '<w:p><w:pPr>' + pPr + '</w:pPr><w:r><w:t>' + t + '</w:t></w:r></w:p>';

test('auto line spacing given as a measure, as Strict files give it, is that many twips over 240', async () => {
  const r = await open(para('One', '<w:spacing w:line="12.95pt" w:lineRule="auto"/>') + para('Two', '<w:spacing w:line="259" w:lineRule="auto"/>'), MAR, '',
    `() => L.D.doc.main.blocks.slice(0, 2).map((p) => +p.pPr.sp.line.toFixed(4))`);
  assert.deepEqual(r, [1.0792, 1.0792]);
});

test('a gutter at the top of the page adds to the top margin', async () => {
  const where = `() => { const pg = L.layout.pages[0]; const b = pg.el.querySelector('.pg-body').getBoundingClientRect(), p = pg.el.getBoundingClientRect(); return [Math.round((b.top - p.top) * 0.75), Math.round((b.left - p.left) * 0.75), Math.round(b.width * 0.75)]; }`;
  // Word 2016: the text starts 108 pt down, 72 pt in, 468 pt wide; with the gutter on the left, 72 down, 108 in, 432 wide
  assert.deepEqual(await open(para('Top'), MAR, '<w:gutterAtTop/>', where), [108, 72, 468]);
  assert.deepEqual(await open(para('Left'), MAR, '', where), [72, 108, 432]);
});
