import test from 'node:test';
import assert from 'node:assert/strict';
import { pathToFileURL } from 'node:url';
import '../../public/common/core.js';
import '../../public/common/xml.js';
import '../../public/common/opc.js';
import '../../public/common/opc-order.js';
import '../../public/common/zip.js';
globalThis.window = globalThis;
await import('../../public/lectern/js/model.js');
await import('../../public/lectern/js/properties.js');
await import(process.env.FRAMES_SOURCE ? pathToFileURL(process.env.FRAMES_SOURCE).href : '../../public/lectern/js/frames.js');
const K = L.opc, N = K.NS, part = 'ppt/slides/slide1.xml';
L.comments = { shapeIdentity: (s, ctx, xml) => xml };
const pkg = async xml => K.open(new Uint8Array(await (await L.zip.write([
  { name: '[Content_Types].xml', data: `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/></Types>` },
  { name: part, data: xml },
])).arrayBuffer()));
const transform = (tag, attrs = '') => `<${tag}${attrs}><a:off x="127000" y="254000"/><a:ext cx="1270000" cy="762000"/></${tag}>`;

test('charts and groups containing charts have no rotation command', () => {
  assert.equal(L.model.canRotate({ type: 'chart' }), false);
  assert.equal(L.model.canRotate({ type: 'group', kids: [{ type: 'chart' }] }), false);
  assert.equal(L.model.canRotate({ type: 'shape' }), true);
  const xml = `<p:graphicFrame xmlns:p="${N.p}" xmlns:a="${N.a}">${transform('p:xfrm')}</p:graphicFrame>`;
  const moved = K.parse(K.setBox(xml, { x: 30, rot: 0, flipH: false, flipV: false })).children[0];
  assert(!['rot', 'flipH', 'flipV'].some(name => moved.hasAttribute(name)));
});

test('a single retained placeholder reads its original horizontal and vertical flips', () => {
  const tree = K.parse(`<am3d:model3d xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:a="${N.a}"><am3d:spPr>${transform('a:xfrm', ' flipH="1" flipV="true"')}</am3d:spPr></am3d:model3d>`);
  const shape = L.frames.placeholder(tree, tree, null, {});
  assert.equal(shape.flipH, true); assert.equal(shape.flipV, true);
});

test('one preview over differently oriented Choice and Fallback shapes moves without overwriting either orientation', async () => {
  const xml = `<mc:AlternateContent xmlns:mc="${N.mc}" xmlns:p="${N.p}" xmlns:a="${N.a}" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><mc:Choice Requires="cx"><p:graphicFrame>${transform('p:xfrm')}<a:graphic><a:graphicData><cx:chart/></a:graphicData></a:graphic></p:graphicFrame></mc:Choice><mc:Fallback><p:pic><p:spPr>${transform('a:xfrm', ' rot="1800000" flipH="1"')}</p:spPr></p:pic></mc:Fallback></mc:AlternateContent>`;
  const p = await pkg(xml), root = K.parse(xml); K.captureAC(root);
  const fallback = root.children[1].children[0];
  const shape = { type: 'image', id: 'preview', x: 10, y: 20, w: 100, h: 60, rot: 30, flipH: true };
  L.frames.attach([shape], fallback, fallback, null, p, part);
  const pres = { slides: [{ shapes: [shape] }], designs: {}, keep: {}, pkg: p };
  await L.frames.seal(pres, new Map());
  shape.x += 12; shape.w *= 1.2;
  const writer = new K.Writer(p), ctx = { writer, part, frames: L.frames.prepare(pres, writer) };
  const saved = K.parse(L.frames.emit(shape, ctx));
  const boxes = saved.children.map(branch => K.getBox(K.raw(branch.children[0])));
  assert.deepEqual(boxes.map(b => [b.rot, b.flipH, b.flipV]), [[0, false, false], [30, true, false]]);
  assert(boxes.every(b => b.x === 22 && b.w === 120));
});

test('replacing a slide extension retains its position among unknown siblings', async () => {
  const ext = (uri, text = '') => `<p:ext uri="${uri}">${text}</p:ext>`;
  const list = value => `<p:extLst>${value}</p:extLst>`;
  const xml = `<p:sld xmlns:p="${N.p}"><p:cSld/>${list(ext('before') + ext('comments', '<p:old/>') + ext('after'))}</p:sld>`;
  const source = await pkg(xml), slide = { keep: {} }; L.properties.slide(slide, source.xml(part), source, part);
  const generated = `<p:sld xmlns:p="${N.p}"><p:cSld/>${list(ext('comments', '<p:current/>'))}</p:sld>`;
  const writer = new K.Writer(source), result = K.parse(L.properties.finishSlide(slide, { writer, part }, generated));
  const extensions = result.children.find(e => e.localName === 'extLst').children;
  assert.deepEqual(extensions.map(e => e.getAttribute('uri')), ['before', 'comments', 'after']);
  assert.equal(extensions[1].children[0].localName, 'current');
});
