// Text in an AutoShape (Word's Add Text, or typing on a selected shape) is saved as Word writes it: the shape stays a
// shape (wps:cNvSpPr without txBox), its text centred as Word starts it; a text box stays a text box. Every shape
// with text used to be written as a text box once edited, and a shape without text could not take any.
// node --test tests/quire/shape-text.test.mjs
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

/* a document holding an ellipse with text (an AutoShape) and a text box, as Word writes them */
const SHAPE = (id, name, prst, txBox, text) => `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251659264 + id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${id * 1500000}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1270000" cy="635000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="${name}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr${txBox ? ' txBox="1"' : ''}/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="${txBox ? 't' : 'ctr'}" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
const BODY = `<w:p>${SHAPE(1, 'Oval 1', 'ellipse', false, 'Inside')}${SHAPE(2, 'Text Box 2', 'rect', true, 'Boxed')}<w:r><w:t>Anchor paragraph.</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>`;

const run = (fn) => page.evaluate(`(${fn.toString()})(${JSON.stringify(BODY)})`);

test('a Word shape with text stays a shape when its text is edited; a text box stays a text box', async () => {
  const sp = await run(async (body) => {
    const K = L.opc, N = K.NS;
    const zip = await L.zip.write([
      { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>` },
      { name: '_rels/.rels', data: `<Relationships xmlns="${N.pkg}"><Relationship Id="rId1" Type="${N.rel}/officeDocument" Target="word/document.xml"/></Relationships>` },
      { name: 'word/document.xml', data: `<w:document xmlns:w="${N.w}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>` },
    ]);
    await L.app.openBytes(new Uint8Array(await zip.arrayBuffer()), 'shapes.docx');
    await new Promise((r) => setTimeout(r, 400));
    const D = L.D, p = D.doc.main.blocks[0];
    L.ed.edit('t', () => { D.touch(p); for (const it of p.runs.filter((r) => r.tb)) { const q = it.tb.blocks[0]; L.O.insertText(D.pos(q, D.plen(q)), ' edited'); } return L.ed.sel; });
    const pk = await K.open(new Uint8Array(await (await L.docx.write(D.doc)).arrayBuffer()));
    return [...pk.text('word/document.xml').matchAll(/<wps:wsp>[\s\S]*?<\/wps:wsp>/g)].map((m) => [(m[0].match(/<wps:cNvSpPr[^>]*>/) || [''])[0], (m[0].match(/<w:t>([^<]*)<\/w:t>/) || [])[1]]);
  });
  assert.deepEqual(sp, [['<wps:cNvSpPr/>', 'Inside edited'], ['<wps:cNvSpPr txBox="1"/>', 'Boxed edited']]);
});

test('Add Text on a new shape: its text centred both ways, saved as a shape with text', async () => {
  const out = await run(async () => {
    const D = L.D, d = D.newDoc();
    d.main.blocks = [D.para([D.text('First paragraph.')]), D.para([D.text('Second paragraph.')])];
    const it = L.drawing.newShape('rect', 120, 80); L.drawing.setWrap(it, 'none');
    it.float.posH = { rel: 'page', off: 300 }; it.float.posV = { rel: 'page', off: 200 };
    d.main.blocks[1].runs.unshift(it);
    L.app.loadDoc(d, 'add text');
    await new Promise((r) => setTimeout(r, 400));
    const wrap = document.querySelector('#pages .flt');
    L.drawing.select(wrap, { p: d.main.blocks[1], o: 0 });
    const label = L.ui.cmds.shapeText.label;
    L.drawing.editText('Hi');
    const pk = await L.opc.open(new Uint8Array(await (await L.docx.write(D.doc)).arrayBuffer()));
    const x = (pk.text('word/document.xml').match(/<wps:wsp>[\s\S]*?<\/wps:wsp>/) || [''])[0];
    return { label, caretInShape: L.ed.sel.f.p === it.tb.blocks[0], objSel: !!L.ed.objSel, cNvSpPr: (x.match(/<wps:cNvSpPr[^>]*>/) || [''])[0], name: (x.match(/<wps:cNvPr[^>]*name="([^"]*)"/) || [])[1], jc: /<w:jc w:val="center"\/>/.test(x), anchor: (x.match(/anchor="(\w+)"/) || [])[1], text: (x.match(/<w:t>([^<]*)<\/w:t>/) || [])[1] };
  });
  assert.equal(out.label, 'Add Te&xt');
  assert.equal(out.caretInShape, true, 'the caret in its text, the shape no longer selected');
  assert.equal(out.objSel, false);
  assert.equal(out.cNvSpPr, '<wps:cNvSpPr/>');
  assert.doesNotMatch(out.name, /Text Box/);
  assert.equal(out.jc, true, 'centred, as Word starts a shape\'s text');
  assert.equal(out.anchor, 'ctr');
  assert.equal(out.text, 'Hi');
});
