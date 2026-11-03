// A row's and a cell's properties keep what Quire doesn't model (w:cnfStyle, w:wAfter, a row's w:tblCellSpacing, cell
// revisions) and a cell's width as the file gives it (a percentage, auto) while the table's grid is unchanged; they were
// rebuilt from the model, every width in twips. A row's exceptions (w:tblPrEx) and the columns' tracked change
// (w:tblGridChange) are kept too; both were dropped.
// node --test tests/quire/table-properties.test.mjs
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
  const cnf = v => `<w:cnfStyle w:val="${v}" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="1" w:evenHBand="0" w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" w:lastRowLastColumn="0"/>`;
  const cell = (pr, text) => `<w:tc><w:tcPr>${pr}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const body = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/><w:gridCol w:w="1000"/>' +
    '<w:tblGridChange w:id="9"><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid></w:tblGridChange></w:tblGrid>' +
    `<w:tr><w:trPr>${cnf('000000100000')}<w:gridAfter w:val="1"/><w:wAfter w:w="1000" w:type="dxa"/><w:cantSplit/><w:tblCellSpacing w:w="20" w:type="dxa"/></w:trPr>` +
    cell(`${cnf('001000100000')}<w:tcW w:w="2000" w:type="pct"/><w:vAlign w:val="center"/>`, 'One') +
    cell('<w:tcW w:w="0" w:type="auto"/><w:cellIns w:id="7" w:author="A" w:date="2026-01-01T00:00:00Z"/>', 'Two') + '</w:tr>' +
    // the second row's exceptions to the table's properties: Word moved its text 54 pt when they were dropped (I19)
    '<w:tr><w:tblPrEx><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:tblCellMar><w:left w:w="1200" w:type="dxa"/></w:tblCellMar></w:tblPrEx>' +
    cell('<w:tcW w:w="2000" w:type="dxa"/>', 'Three') + cell('<w:tcW w:w="3000" w:type="dxa"/>', 'Four') + '</w:tr></w:tbl><w:p/>';
  const original = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body>${body}<w:sectPr/></w:body></w:document>` },
  ]);
  const { doc } = await L.docx.read(new Uint8Array(await original.arrayBuffer()));
  const table = doc.main.blocks[0];
  if (kind === 'resized') {
    // a column dragged wider, as the ruler and the table's edges do it: the grid changes, the cells don't
    table.grid = [table.grid[0] + 36, table.grid[1] - 36, table.grid[2]];
    table.rows[0].trPr = { ...table.rows[0].trPr, cantSplit: false };
  }
  const saved = await K.open(new Uint8Array(await (await L.docx.write(doc)).arrayBuffer()));
  const root = saved.xml('word/document.xml'), w = (el, tag) => Array.from(el.getElementsByTagNameNS(N.w, tag));
  const children = el => el ? Array.from(el.children).map(c => c.localName + (c.getAttributeNS(N.w, 'type') ? ':' + c.getAttributeNS(N.w, 'type') : '')) : [];
  return { row: children(w(root, 'trPr')[0]), cells: w(root, 'tcPr').slice(0, 2).map(children), exceptions: w(root, 'tblPrEx').map(children),
    grid: children(w(root, 'tblGrid')[0]), losses: (doc.losses || []).filter(e => e.id.startsWith('grid-change:')).length };
}
const run = async kind => {
  const result = await page.evaluate(`(${save.toString()})(${JSON.stringify(kind)})`);
  assert.deepEqual(page.errors, [], 'no browser exceptions, console errors or CSP refusals');
  return result;
};

test('an unedited save keeps what a row and a cell have, widths as the file gives them', async () => {
  const got = await run('unedited');
  assert.deepEqual(got.row, ['cnfStyle', 'gridAfter', 'wAfter:dxa', 'cantSplit', 'tblCellSpacing:dxa']);
  assert.deepEqual(got.cells, [['cnfStyle', 'tcW:pct', 'vAlign'], ['tcW:auto', 'cellIns']]);
  assert.deepEqual(got.exceptions, [['shd', 'tblCellMar']], 'the row\'s exceptions');
  assert.deepEqual(got.grid, ['gridCol', 'gridCol', 'gridCol', 'tblGridChange'], 'the columns\' tracked change');
});

test('a resized column gives the cells new widths; a changed row property goes, the rest stays', async () => {
  const got = await run('resized');
  assert.deepEqual(got.row, ['cnfStyle', 'gridAfter', 'wAfter:dxa', 'tblCellSpacing:dxa']);
  assert.deepEqual(got.cells, [['cnfStyle', 'tcW:dxa', 'vAlign'], ['tcW:dxa', 'cellIns']]);
  assert.deepEqual(got.exceptions, [['shd', 'tblCellMar']]);
  assert.deepEqual([got.grid, got.losses], [['gridCol', 'gridCol', 'gridCol'], 1], 'the columns\' tracked change goes with a record, as the columns changed');
});
