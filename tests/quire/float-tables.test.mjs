// A floating table (w:tblpPr) is laid out where its anchors put it, with the text beside it, and one longer than the
// page goes on at the top of the next; Table Properties' Text wrapping makes a table floating and puts it back. Every
// floating table used to be drawn in the text, where it stood.
// node --test tests/quire/float-tables.test.mjs
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

/* opens a Letter page with 1" margins (compatibility mode 15: the border stands at the offset) holding body; returns
   the tables' boxes, the page's and the first line beside, in px at 100% (from the page's corner) */
async function lay(body) {
  const K = L.opc, N = K.NS;
  const zip = await L.zip.write([
    { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>` },
    { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="word/document.xml"/></Relationships>` },
    { name: 'word/_rels/document.xml.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/settings" Target="settings.xml"/></Relationships>` },
    { name: 'word/settings.xml', data: `<w:settings xmlns:w="${N.w}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>` },
    { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>` },
  ]);
  await L.app.openBytes(new Uint8Array(await zip.arrayBuffer()), 'float.docx');
  L.app.zoomTo('100%');
  await new Promise((r) => setTimeout(r, 400));
  const pages = L.layout.pages.map((p) => p.el.getBoundingClientRect());
  const at = (r) => { const i = pages.findIndex((p) => r.top >= p.top - 1 && r.top < p.bottom); const p = pages[i]; return { page: i + 1, x: Math.round(r.left - p.left), y: Math.round(r.top - p.top), w: Math.round(r.width), h: Math.round(r.height) }; };
  const tables = [...document.querySelectorAll('#pages table.tbl')].map((e) => at(e.getBoundingClientRect()));
  /* the first line of the text after the table: where it starts */
  const after = [...document.querySelectorAll('#pages .p')].find((e) => /Beside/.test(e.textContent));
  const rg = document.createRange();
  if (after) { const tn = [...after.querySelectorAll('*')].flatMap((e) => [...e.childNodes]).find((n) => n.nodeType === 3 && /Beside/.test(n.data)); rg.setStart(tn, 0); rg.setEnd(tn, 1); }
  return { pages: pages.length, tables, beside: after ? at(rg.getClientRects()[0]) : null };
}

const P = (t) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const TEXT = P('Before.') + '%TABLE%' + P('Beside the table. ' + 'Text wraps around the floating table as it does in Word. '.repeat(8));
function table(tblpPr, rows = 3) {
  const cell = (r, c) => `<w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/></w:tcPr>${P(`R${r}C${c}`)}</w:tc>`;
  const rowsXML = Array.from({ length: rows }, (_, r) => `<w:tr>${cell(r + 1, 1)}${cell(r + 1, 2)}</w:tr>`).join('');
  return `<w:tbl><w:tblPr>${tblpPr}<w:tblW w:w="2880" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>${rowsXML}</w:tbl>`;
}
const run = (body) => page.evaluate(`(${lay.toString()})(${JSON.stringify(body)})`);

test('a text-anchored table stands at the margin, the text beside it', async () => {
  const r = await run(TEXT.replace('%TABLE%', table('<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="1"/>')));
  const [t] = r.tables;
  assert.equal(t.x, 96, 'its border at the left margin (1")');
  assert.ok(r.beside.y >= t.y - 2 && r.beside.y < t.y + 20, `the next paragraph starts beside its top (${r.beside.y} by ${t.y})`);
  assert.ok(r.beside.x >= t.x + t.w + 10, `to its right, past rightFromText's 0.125" (${r.beside.x})`);
});

test('aligned right of the margin, the text goes on its left', async () => {
  const r = await run(TEXT.replace('%TABLE%', table('<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="1"/>')));
  const [t] = r.tables;
  assert.ok(Math.abs(t.x + t.w - (816 - 96)) <= 1, `its right edge at the right margin (${t.x + t.w})`);
  assert.equal(r.beside.x, 96);
  assert.ok(r.beside.y < t.y + 20);
});

test('page-anchored: 3" from the page\'s top, centred on the page, whatever the text before it', async () => {
  const r = await run(P('Text before. '.repeat(30)) + table('<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="page" w:horzAnchor="page" w:tblpXSpec="center" w:tblpY="4320"/>') + P('Beside it.'));
  const [t] = r.tables;
  assert.equal(t.y, 288);
  assert.ok(Math.abs(t.x - (816 - t.w) / 2) <= 1, `centred (${t.x})`);
});

test('a floating table longer than the page goes on at the top of the next, the text after its last part', async () => {
  const r = await run(TEXT.replace('%TABLE%', table('<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="1"/>', 40)));
  assert.equal(r.tables.length, 2, 'two parts: ' + JSON.stringify(r.tables));
  assert.deepEqual(r.tables.map((t) => [t.page, t.x]), [[1, 96], [2, 96]]);
  assert.ok(Math.abs(r.tables[1].y - 96) <= 1, `the second at the top of page 2's body (${r.tables[1].y})`);
  assert.equal(r.beside.page, 2);
});

test('Table Properties: Text wrapping Around makes it floating as Word does, None puts it back', async () => {
  const xml = await page.evaluate(`(async () => {
    const D = L.D, E = L.ed, t = D.doc.main.blocks.find((b) => b.t === 'tbl');
    E.setSel(D.pos(t.rows[0].cells[0].blocks[0], 0));
    const pick = async (v) => { L.dlg.tableProps(); await new Promise((r) => setTimeout(r, 200)); [...document.querySelectorAll('.dlg input[name=tp-wr]')][v === 'none' ? 0 : 1].click(); [...document.querySelectorAll('.dlg .dlg-foot button')].find((b) => b.textContent === 'OK').click(); await new Promise((r) => setTimeout(r, 300)); };
    const save = async () => { const K = L.opc; const pk = await K.open(new Uint8Array(await (await L.docx.write(D.doc)).arrayBuffer())); return (pk.text('word/document.xml').match(/<w:tblpPr[^>]*>/) || [''])[0]; };
    const out = [];
    await pick('none'); out.push(await save());
    await pick('around'); out.push(await save());
    return out;
  })()`);
  assert.equal(xml[0], '', 'None: no w:tblpPr');
  assert.equal(xml[1], '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:tblpY="1"/>');
});
