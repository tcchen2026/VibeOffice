// Drawings and their room on the page, against Word 2016's prints (the Office VM, 2026-11-07: documents like the ones
// built here, and the corpus's tdf151518_SmartArtTextLocation.docx and complexity - part 1.docx; the Quire-vs-Word
// pictures of tools/quire-pictures.mjs found them):
// - an object's wp:effectExtent (room for a shadow, a glow, a turned picture's corners) is a margin around it in its
//   line, and adds to a floating object's distance from the text: Word set "X", a 72 × 36 pt picture with 36 pt all
//   round, "Y" on one line 108 pt tall, the picture 36 pt in from either side;
// - a picture in line with text wider than its line sticks out past the right edge on its one line;
// - a text box that fits its text is fitted again when the page is laid out, shrunk as well as grown;
// - a line holding nothing but anchors (an empty paragraph) is moved below an object text can't go beside, as any line;
// - each SmartArt diagram is drawn from its own drawing part (its data's dataModelExt), with its text's bullets;
// - pct25 shading is a quarter of the pattern colour over the fill: light grey, black text, for auto on auto;
// - STYLEREF finds a style by the name Word shows with its aliases ("Product Head,ph") or by an alias, and the first
//   paragraph in the style on the page gives its text even when it is empty;
// - a WMF whose placeable header has empty bounds is sized by its window;
// - VML freeforms and polylines are drawn from their paths and points, stacked by every step of their z-index, and a
//   v:rect's textpath is not WordArt (Word draws its text box).
// node --test tests/quire/drawn-as-word.test.mjs
import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from '../../tools/ooxml/cdp.mjs';

const corpora = process.env.VO_CORPORA || '/work/corpora';
const diagrams = path.join(corpora, 'word', 'libreoffice__65c10c3ce0a3__tdf151518_SmartArtTextLocation.docx');
let page, stop;
before(async () => {
  stop = await ensureServer(process.env.BASE);
  page = await openPage('quire', { persistence: false, automated: true });
});
after(async () => {
  try { await page?.close(); } finally { stop?.(); }
});

const E = 12700;
const PIC = (w, h, ee, anchor) => {
  const [l, t, r, b] = ee.map((v) => v * E), g = `<wp:extent cx="${w * E}" cy="${h * E}"/><wp:effectExtent l="${l}" t="${t}" r="${r}" b="${b}"/>${anchor ? anchor[1] : ''}<wp:docPr id="1" name="Picture 1"/><wp:cNvGraphicFramePr/>`
    + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="red.png"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="rIdP"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * E}" cy="${h * E}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>`;
  if (!anchor) return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">${g}</wp:inline></w:drawing></w:r>`;
  return `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251658241" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>${anchor[0]}${g}</wp:anchor></w:drawing></w:r>`;
};
const at = (x, y, wrap = '<wp:wrapTopAndBottom/>') => [`<wp:positionH relativeFrom="column"><wp:posOffset>${x * E}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${y * E}</wp:posOffset></wp:positionV>`, wrap];
const p = (...runs) => '<w:p>' + runs.map((r) => (r.startsWith('<w:r>') ? r : `<w:r><w:t xml:space="preserve">${r}</w:t></w:r>`)).join('') + '</w:p>';

/** opens a document (Arial 12 pt, single spacing, 1 in margins; rIdP a red picture; styles and footer as given) and
    returns where things stand on the first page (pt from its top left): each picture, each line's text, each table
    cell's background and text colour, the footer's text on each page */
async function open(body, { styles = '', footer = '' } = {}) {
  const r = await page.evaluate(`(async () => {
    const K = L.opc, N = K.NS, part = 'word/document.xml', T = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
    const c = new OffscreenCanvas(8, 8), g = c.getContext('2d'); g.fillStyle = '#FF0000'; g.fillRect(0, 0, 8, 8);
    const png = new Uint8Array(await (await c.convertToBlob({ type: 'image/png' })).arrayBuffer());
    const footer = ${JSON.stringify(footer)};
    const files = [
      { name: '[Content_Types].xml', data: '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/' + part + '" ContentType="' + T + 'document.main+xml"/><Override PartName="/word/styles.xml" ContentType="' + T + 'styles+xml"/><Override PartName="/word/footer1.xml" ContentType="' + T + 'footer+xml"/></Types>' },
      { name: '_rels/.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/officeDocument" Target="' + part + '"/></Relationships>' },
      { name: 'word/_rels/document.xml.rels', data: '<Relationships xmlns="' + N.pkg + '"><Relationship Id="rId1" Type="' + N.rel + '/styles" Target="styles.xml"/><Relationship Id="rIdP" Type="' + N.rel + '/image" Target="media/red.png"/><Relationship Id="rIdF" Type="' + N.rel + '/footer" Target="footer1.xml"/></Relationships>' },
      { name: 'word/styles.xml', data: '<w:styles xmlns:w="' + N.w + '"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' + ${JSON.stringify(styles)} + '</w:styles>' },
      { name: 'word/media/red.png', data: png },
      { name: 'word/footer1.xml', data: '<w:ftr xmlns:w="' + N.w + '"><w:p>' + footer + '</w:p></w:ftr>' },
      { name: part, data: '<w:document xmlns:w="' + N.w + '" xmlns:r="' + N.r + '" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body>' + ${JSON.stringify(body)} + '<w:sectPr>' + (footer ? '<w:footerReference w:type="default" r:id="rIdF"/>' : '') + '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>' },
    ];
    await L.app.openFile(new File([await L.zip.write(files)], 'drawn.docx'));
    L.layout.setZoom(1);
    await new Promise((r) => setTimeout(r, 400)); await document.fonts.ready;
    const pg = L.layout.pages[0].el.getBoundingClientRect(), pt = (v, o) => Math.round((v - o) * 0.75 * 10) / 10;
    const box = (q) => [pt(q.left, pg.left), pt(q.top, pg.top), pt(q.right, pg.left), pt(q.bottom, pg.top)];
    const pics = [...L.layout.pages[0].el.querySelectorAll('.img img')].map((e) => box(e.getBoundingClientRect()));
    const shapes = [...L.layout.pages[0].el.querySelectorAll('.shp')].map((e) => box(e.getBoundingClientRect()));
    /* each floating shape's outline (its path's corners, page pt) and its place in the stack */
    const outlines = [...L.layout.pages[0].el.querySelectorAll('.flt')].filter((w) => w.querySelector('svg.geo path')).map((w) => {
      const path = w.querySelector('svg.geo path'), m = path.getScreenCTM(), d = path.getAttribute('d');
      const pts = [...d.matchAll(/(-?[\\d.]+),(-?[\\d.]+)/g)].map(([, x, y]) => { const q = new DOMPoint(+x, +y).matrixTransform(m); return [pt(q.x, pg.left), pt(q.y, pg.top)]; });
      return { name: w._item.name || '', z: +getComputedStyle(w).zIndex, pts, cmds: d.replace(/[^A-Za-z]/g, '') };
    });
    const text = {};
    for (const s of L.layout.pages[0].el.querySelectorAll('.pg-body .pc > span')) for (const w of s.textContent.trim().split(/\\s+/)) if (w && !(w in text)) {
      const rg = document.createRange(), t = s.firstChild; if (!t || t.nodeType !== 3) continue;
      const i = t.data.indexOf(w); rg.setStart(t, i); rg.setEnd(t, i + w.length); text[w] = box(rg.getBoundingClientRect());
    }
    const cells = [...L.layout.pages[0].el.querySelectorAll('td')].map((td) => [getComputedStyle(td).backgroundColor, getComputedStyle(td.querySelector('.pc > span') || td).color]);
    const footers = L.layout.pages.map((x) => (x.el.querySelector('.pg-ftr') || {}).textContent || '');
    return { pics, shapes, outlines, text, cells, footers };
  })()`);
  assert.deepEqual(page.errors, []);
  return r;
}
const near = (got, want, what) => assert.ok(Math.abs(got - want) <= 1.5, `${what}: ${got} pt, Word ${want}`);

test('a picture in line with text has the room for its effects around it', async () => {
  const r = await open(p('Above') + p('X', PIC(72, 36, [36, 36, 36, 36]), 'Y') + p('Below'));
  const [x, y] = r.pics[0];
  // Word 2016: the picture at 116, 121.5 (X is 8 pt wide); Y at 224; the line 108 pt tall, its baseline at 194
  near(x, 116, 'the picture\'s left: 36 pt after X');
  near(y, 121.5, 'its top: 36 pt below the line above');
  near(r.text.Y[0], 224, 'Y: 36 pt after the picture');
  near(r.text.Below[1], 196.5, 'the next line: below the 36 pt under the picture');
});

test('a picture in line with text wider than its line sticks out past the edge, on the one line', async () => {
  // Word 2016: a 500 pt picture in a 468 pt column, alone in its paragraph, left or centred: it starts at the left
  // indent, and the next line starts at its bottom (121.8 pt); it took a second, empty line in Quire
  const r = await open(p('Above') + p(PIC(500, 36, [0, 0, 0, 0])) + p('Next') + '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' + PIC(500, 36, [0, 0, 0, 0]) + '</w:p>' + p('Last'));
  assert.deepEqual(r.pics.map((q) => Math.round(q[0])), [72, 72]);
  near(r.text.Next[1], 121.8, 'the line after the first');
  near(r.text.Last[1], 171.8, 'the line after the centred one');
});

test('a floating picture keeps the text its distance plus the room for its effects away', async () => {
  // top and bottom, 24 pt all round: drawn where its position says (144, 108); the text resumes 24 pt below it
  let r = await open(p('Para', PIC(144, 72, [24, 24, 24, 24], at(72, 36))) + p('Line'));
  assert.deepEqual(r.pics[0].slice(0, 2), [144, 108]);
  near(r.text.Para[1], 204.5, 'its paragraph\'s line, moved below it');
  // square, 36 pt on the right: the text beside starts 36 pt after it
  r = await open(p(PIC(144, 72, [0, 0, 36, 0], at(0, 0, '<wp:wrapSquare wrapText="bothSides"/>')), 'Text beside the picture '.repeat(12)));
  near(r.text.Text[0], 252, 'the text beside');
});

test('a text box that fits its text is as tall as its text, shrunk from the size saved', async () => {
  // Word 2016 draws a box saved 200 pt tall around two lines 34.5 pt tall (two 13.8 pt lines and its 3.6 pt insets)
  const box = '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' + at(0, 20)[0]
    + `<wp:extent cx="${200 * E}" cy="${200 * E}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="Text Box 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">`
    + `<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${200 * E}" cy="${200 * E}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>`
    + '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Two lines of text in a box that is meant to fit its text exactly</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:spAutoFit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
  const r = await open('<w:p>' + box + '<w:r><w:t>Anchor</w:t></w:r></w:p>');
  const [, top, , bottom] = r.shapes[0];
  near(top, 92, 'its top');
  near(bottom - top, 34.5, 'its height');
});

test('an empty line goes below an object text can\'t go beside', async () => {
  // the picture from 92 to 192 pt; the empty paragraph's line (85.8 to 99.6) would cross it: Word moves it below, so
  // Beta starts a line further down (tdf151518_SmartArtTextLocation.docx's empty paragraphs between its diagrams)
  const r = await open(p('Alpha', PIC(144, 100, [0, 0, 0, 0], at(0, 20))) + p() + p('Beta'));
  near(r.text.Beta[1], 205.8, 'Beta');
});

test('pct25 shading is a quarter of the pattern colour over the fill', async () => {
  const r = await open('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="pct25" w:color="auto" w:fill="auto"/></w:tcPr>' + p('Shaded') + '</w:tc></w:tr></w:tbl>' + p());
  assert.deepEqual(r.cells[0], ['rgb(191, 191, 191)', 'rgb(0, 0, 0)']);
});

test('STYLEREF: a style named with its aliases or by one; the first paragraph in it on the page, even empty', async () => {
  const fld = (instr) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>stale</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const t = (s) => `<w:r><w:t xml:space="preserve">${s}</w:t></w:r>`;
  const head = (s) => `<w:p><w:pPr><w:pStyle w:val="ProductHead"/></w:pPr>${s ? t(s) : ''}</w:p>`;
  const r = await open(head('') + head('Alpha') + p('Body') + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + p('More') + head('Beta'), {
    styles: '<w:style w:type="paragraph" w:customStyle="1" w:styleId="ProductHead"><w:name w:val="Product Head"/><w:aliases w:val="ph"/><w:basedOn w:val="Normal"/></w:style>',
    footer: t('A[') + fld('STYLEREF "Product Head,ph"') + t('] B[') + fld('STYLEREF ph') + t(']'),
  });
  // Word 2016: nothing on the first page (its first Product Head paragraph is empty), Beta on the second
  assert.deepEqual(r.footers, ['A[] B[]', 'A[Beta] B[Beta]']);
});

test('each SmartArt diagram is drawn from its own drawing, its text with its bullets', async (t) => {
  if (!fs.existsSync(diagrams)) return t.skip('Fetch pinned corpus: ' + diagrams);
  const r = await page.evaluate(`(async () => {
    await L.app.openFile(new File([Uint8Array.from(atob(${JSON.stringify(fs.existsSync(diagrams) ? fs.readFileSync(diagrams).toString('base64') : '')}), (c) => c.charCodeAt(0))], 'd.docx'));
    const out = [];
    const text = (k) => (k.tb ? k.tb.blocks.map((b) => b.runs.map((x) => x.text || '').join('')).join('|') : '');
    for (const q of L.D.allParas(L.D.doc)) for (const it of q.runs || []) if (it.diagram) out.push(it.kids.map((k) => k.geom + ' ' + Math.round(k.w)).join(', ') + ': ' + it.kids.map(text).filter(Boolean).join(' '));
    return out;
  })()`);
  assert.deepEqual(page.errors, []);
  // Target List, Nested Target, Stacked Venn, Grouped List: four drawings (all four were drawn as the first)
  assert.equal(r.length, 4);
  assert.equal(new Set(r).size, 4, 'four different drawings');
  assert.ok(/^(ellipse \d+, ){3}ellipse \d+: Cat Dog Pig Cow$/.test(r[2]), 'Stacked Venn: four circles: ' + r[2]);
  assert.ok(r[0].includes('•Cat|•Dog'), 'Target List\'s bulleted children: ' + r[0]);
});

test('a WMF whose placeable bounds are empty is drawn to its window', async (t) => {
  const file = path.join(corpora, 'word', 'openxml-sdk__8a305280c46e__complexity - part 1.docx');
  if (!fs.existsSync(file)) return t.skip('Fetch pinned corpus: ' + file);
  // its header says 1735, 32767 to 2259, 32767; its window is 4192 × 136 (Word draws the black banner from it)
  const r = await page.evaluate(`(async () => {
    await L.app.openFile(new File([Uint8Array.from(atob(${JSON.stringify(fs.existsSync(file) ? fs.readFileSync(file).toString('base64') : '')}), (c) => c.charCodeAt(0))], 'c.docx'));
    const im = document.querySelector('.pg .img img');
    await (im.decode ? im.decode().catch(() => {}) : null);
    return [im.naturalWidth, im.naturalHeight];
  })()`);
  assert.deepEqual(page.errors, []);
  assert.ok(Math.abs(r[0] / r[1] - 4192 / 136) < 1, 'drawn at its window\'s shape: ' + r);
});

test('VML freeforms and polylines are drawn from their paths and points, stacked by their z-index', async () => {
  // tdf159158_zOrder_zIndexMax.docx's two stars (a path of ten corners in a 1924685 × 1638935 coordsize) and
  // tdf112450_vml_polyline.docx's polylines: Word 2016 draws them so; Quire drew rectangles, the higher star below
  const star = (name, x, z, colour) => `<w:r><w:pict><v:shape id="${name}" style="position:absolute;margin-left:${x}pt;margin-top:41pt;width:151.55pt;height:129.05pt;z-index:${z}" coordsize="1924685,1638935"`
    + ` path="m2,626016r735166,4l962343,r227174,626020l1924683,626016r-594765,386897l1557102,1638931,962343,1252026,367583,1638931,594767,1012913,2,626016xe" fillcolor="${colour}"/></w:pict></w:r>`;
  const poly = '<w:r><w:pict><v:polyline id="Polyline" style="position:absolute;z-index:3" points="76.05pt,77.25pt,143.25pt,94.65pt,229.05pt,40.05pt,271.05pt,53.85pt" filled="f" strokecolor="#9acd32" strokeweight="1.5pt"/></w:pict></w:r>';
  const r = await open('<w:p>' + star('Yellow', 66, 2147483647, 'yellow') + star('Blue', 0, 2147483646, 'blue') + poly + '</w:p>');
  const [yellow, blue, line] = r.outlines;
  assert.equal(yellow.cmds, 'MLLLLLLLLLLZ', 'the star\'s ten corners: ' + yellow.cmds);
  // the top point: 962343 of 1924685 across the 151.55 pt, at the top of the shape (72 + 66, 72 + 41)
  near(yellow.pts[2][0], 72 + 66 + 151.55 * 962343 / 1924685, 'the top point across');
  near(yellow.pts[2][1], 113, 'the top point down');
  assert.ok(yellow.z > blue.z, `the higher z-index on top: ${yellow.z} over ${blue.z}`);
  assert.equal(line.cmds, 'MLLL', 'the polyline\'s four points, open');
  assert.deepEqual(line.pts.map((q) => q.map(Math.round)), [[148, 149], [215, 167], [301, 112], [343, 126]]);
});

test('a v:rect with a textpath is a text box, not WordArt', async () => {
  // drawinglayer-pic-pos.docx: Word 2016 draws the rectangle with its text box's text, the textpath not at all
  const rect = '<w:r><w:pict><v:rect id="r" style="position:absolute;margin-left:0;margin-top:9pt;width:300pt;height:120pt;z-index:2" strokeweight="1.41mm">'
    + '<v:path textpathok="t"/><v:textpath on="t" style="font-family:&quot;Times New Roman&quot;" fitshape="t" string="The Space Race"/>'
    + '<v:textbox><w:txbxContent><w:p><w:r><w:t>Plain words</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r>';
  await open('<w:p>' + rect + '</w:p>');
  const r = await page.evaluate(`(() => { const s = L.layout.pages[0].el.querySelector('.flt .shp'); return { wordart: !!s.querySelector('svg.wa'), text: (s.querySelector('.tbx') || {}).textContent || '' }; })()`);
  assert.deepEqual(r, { wordart: false, text: 'Plain words' });
});
