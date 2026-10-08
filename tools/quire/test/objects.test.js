const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../../public/common/core.js');
require('../../../public/common/xml.js');
require('../../../public/common/opc.js');
require('../../../public/common/opc-order.js');
require('../../../public/quire/js/dmodel.js');
require('../../../public/quire/js/preserve.js');
require('../../../public/quire/js/preserve-drawing.js');
const { D, preserve: P, opc: K } = L;
const part = 'word/document.xml';
const xml = '<w:drawing xmlns:w="' + K.NS.w + '" xmlns:wp="' + K.KNOWN_NS.wp + '" xmlns:a="' + K.NS.a + '"><wp:inline><wp:extent cx="1270000" cy="635000"/><wp:docPr id="42" name="Diagram"/><a:graphic><a:graphicData/></a:graphic></wp:inline></w:drawing>';
function fixture() {
  const doc = D.newDoc(); doc.keep = {}; D.doc = doc;
  const object = P.keepObject({ t: 'group', w: 100, h: 50, kids: [{ t: 'shape', x: 10, y: 5, w: 20, h: 10, fill: { t: 'solid', c: '#123456' } }] }, K.parse(xml), { part }, doc, 'SmartArt');
  doc.main.blocks = [D.para([object])]; P.finishObjects(doc);
  return { doc, object };
}
function write(doc, object, name = 'opaque') {
  const writer = new K.Writer(null, { doc }), ctx = { writer, part, objects: P.prepareObjects(doc, writer) };
  return P.objectXML(object, name, ctx);
}

test('opaque group resize supports both editor scaling and an extent edit without replacing its payload', () => {
  for (const scaleChildren of [false, true]) {
    const { doc, object } = fixture(); object.w = 200;
    if (scaleChildren) { object.kids[0].x *= 2; object.kids[0].w *= 2; }
    assert.match(write(doc, object), /cx="2540000"/);
    assert.equal(doc.losses?.length || 0, 0);
    object.kids[0].fill.c = '#ABCDEF';
    assert.equal(write(doc, object), null); assert.equal(doc.losses[0].action, 'conversion');
  }
});

test('several items from one choice emit once, and deleting one never resurrects its old content', () => {
  const doc = D.newDoc(); doc.keep = {};
  const tree = K.parse(`<w:p xmlns:w="${K.NS.w}" xmlns:mc="${K.NS.mc}"><mc:AlternateContent><mc:Choice Requires="w"><w:r><w:t>A</w:t></w:r><w:r><w:t>B</w:t></w:r></mc:Choice><mc:Fallback/></mc:AlternateContent></w:p>`);
  K.captureAC(tree);
  const p = D.para(), choice = tree.children[0].children[0], ac = P.readAlternates(tree, p.runs, { part }, doc, 'inline');
  for (const [i, el] of [...choice.children].entries()) { p.runs.push(D.text(i ? 'B' : 'A')); ac.add(el, i); }
  ac.finish(); doc.main.blocks = [p]; P.finishObjects(doc);
  const writer = new K.Writer(null, { doc }), ctx = { writer, part, objects: P.prepareObjects(doc, writer) };
  assert.match(P.objectXML(p.runs[0], 'ac', ctx), /AlternateContent/);
  assert.equal(P.objectXML(p.runs[1], 'ac', ctx), '');
  p.runs.pop(); assert.equal(write(doc, p.runs[0], 'ac'), null);
  assert.ok(doc.losses.some(e => e.action === 'conversion'));
});

test('copying VML remaps shape, OLE target and shape-type references together', () => {
  const source = K.fragment(K.parse(`<w:object xmlns:w="${K.NS.w}" xmlns:v="${K.KNOWN_NS.v}" xmlns:o="${K.KNOWN_NS.o}"><v:shapetype id="_x0000_t75"/><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:100pt;height:50pt"/><o:OLEObject ShapeID="_x0000_i1025"/><w:control w:name="TextBox1" w:shapeid="_x0000_i1025"/></w:object>`), { part });
  const copied = K.duplicate(source), writer = new K.Writer(null);
  const a = writer.emit(source, part), b = writer.emit(copied, part);
  const shape = x => /<v:shape id="([^"]+)"/.exec(x)[1];
  assert.notEqual(shape(a), shape(b)); assert.ok(b.includes(`ShapeID="${shape(b)}"`));
  const type = /<v:shapetype id="([^"]+)"/.exec(b)[1];
  assert.ok(b.includes(`type="#${type}"`)); assert.notEqual(type, '_x0000_t75');
  assert.match(a, /w:name="TextBox1"/); assert.match(b, /w:name="TextBox2"/);
});

test('picture crop and border edits retain recolouring, effects and extension children', () => {
  const doc = D.newDoc(); doc.keep = {}; D.doc = doc;
  const source = xml.replace('<a:graphicData/>', `<a:graphicData><pic:pic xmlns:pic="${K.KNOWN_NS.pic}"><pic:nvPicPr><pic:cNvPr id="8" name="Original"/></pic:nvPicPr><pic:blipFill><a:blip><a:duotone><a:srgbClr val="001122"/><a:srgbClr val="FFFFFF"/></a:duotone><a:extLst><a:ext uri="retained"/></a:extLst></a:blip><a:srcRect l="10000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:effectLst><a:glow rad="20000"><a:srgbClr val="AA0000"/></a:glow><a:softEdge rad="30000"/></a:effectLst><a:sp3d z="0"/></pic:spPr></pic:pic></a:graphicData>`);
  const it = P.keepObject({ t: 'img', w: 100, h: 50, crop: { l: 0.1 } }, K.parse(source), { part }, doc, 'Picture');
  doc.main.blocks = [D.para([it])]; P.finishObjects(doc);
  it.crop.l = 0.2; it.border = { val: 'single', sz: 1.5, color: '123456' };
  const writer = new K.Writer(null, { doc }), ctx = { writer, part, objects: P.prepareObjects(doc, writer), drawingXML: () => source.replace('l="10000"', 'l="20000"').replace('<a:effectLst>', '<a:ln w="19050"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln><a:effectLst>') };
  const result = P.objectXML(it, 'opaque', ctx);
  for (const expected of ['l="20000"', 'w="19050"', 'a:duotone', 'uri="retained"', 'a:glow', 'a:softEdge', 'a:sp3d']) assert.ok(result.includes(expected), expected);
  assert.ok(result.indexOf('<a:ln ') < result.indexOf('<a:effectLst'));
  assert.equal(doc.losses?.length || 0, 0);
  it.media = 'a different picture'; assert.equal(write(doc, it), null);
});

test('changing a shadow retains sibling effects in the choice and updates the VML fallback', () => {
  const doc = D.newDoc(); doc.keep = {};
  const drawing = xml.replace('<a:graphicData/>', `<a:graphicData><wps:wsp xmlns:wps="${K.KNOWN_NS.wps}"><wps:spPr><a:effectLst><a:glow rad="20000"><a:srgbClr val="AA0000"/></a:glow><a:outerShdw dist="12700"><a:srgbClr val="000000"/></a:outerShdw><a:reflection dist="30000"/></a:effectLst><a:scene3d/></wps:spPr></wps:wsp></a:graphicData>`);
  const original = `<mc:AlternateContent xmlns:mc="${K.NS.mc}" xmlns:wps="${K.KNOWN_NS.wps}" xmlns:v="${K.KNOWN_NS.v}" xmlns:w="${K.NS.w}"><mc:Choice Requires="wps">${drawing}</mc:Choice><mc:Fallback><w:pict><v:shape id="Fallback" style="width:100pt;height:50pt"><v:shadow on="t" offset="1pt,1pt"/><v:extrusion/></v:shape></w:pict></mc:Fallback></mc:AlternateContent>`;
  const it = P.keepObject({ t: 'shape', w: 100, h: 50, shadow: { dx: 1, dy: 1 } }, K.parse(original), { part }, doc, 'Shape');
  doc.main.blocks = [D.para([it])]; P.finishObjects(doc);
  it.shadow = { dx: 4, dy: 3, c: '#123456' };
  const writer = new K.Writer(null, { doc }), ctx = { writer, part, objects: P.prepareObjects(doc, writer), drawingXML: () => drawing.replace('dist="12700"', 'dist="63500"') };
  const result = P.objectXML(it, 'opaque', ctx);
  for (const expected of ['dist="63500"', 'offset="4pt,3pt"', 'a:glow', 'a:reflection', 'a:scene3d', 'v:extrusion']) assert.ok(result.includes(expected), expected);
  assert.equal(doc.losses?.length || 0, 0);
});

test('editing a property inside compatibility alternatives updates each branch without duplicating it outside', () => {
  const old = '<a:effectLst><a:glow rad="20000"/><a:outerShdw dist="12700"/></a:effectLst>';
  const frame = effects => xml.replace('<a:graphicData/>', `<a:graphicData><wps:wsp xmlns:wps="${K.KNOWN_NS.wps}"><wps:spPr>${effects}<a:sp3d z="0"/></wps:spPr></wps:wsp></a:graphicData>`);
  const source = frame(`<mc:AlternateContent xmlns:mc="${K.NS.mc}" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><mc:Choice Requires="a14">${old}</mc:Choice><mc:Fallback>${old}</mc:Fallback></mc:AlternateContent>`);
  const result = P.mergeDrawing(source, { t: 'shape', shadow: { dx: 5 } }, { shadow: 'null' }, { drawingXML: () => frame(old.replace('dist="12700"', 'dist="63500"')) });
  assert.equal((result.match(/dist="63500"/g) || []).length, 2);
  assert.equal((result.match(/<a:effectLst/g) || []).length, 2);
  assert.equal((result.match(/<a:glow/g) || []).length, 2);
  assert.match(result, /a:sp3d/);
});
