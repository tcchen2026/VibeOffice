import test from 'node:test';
import assert from 'node:assert/strict';
import '../../public/common/core.js';
import '../../public/common/xml.js';
import '../../public/common/opc.js';
import '../../public/common/opc-order.js';
import '../../public/common/zip.js';
globalThis.window = globalThis;
await import('../../public/lectern/js/model.js');
await import('../../public/lectern/js/properties.js');

const K = L.opc, N = K.NS, part = 'ppt/slides/slide1.xml';
const ma = 'http://schemas.microsoft.com/office/mac/drawingml/2008/main';
const pdfPart = 'ppt/media/image1.pdf', pngPart = 'ppt/media/image2.png';
// Media is opaque to this writer: both payloads must survive byte for byte.
const pdf = new TextEncoder().encode('%PDF-1.4\n% Synthetic alternate picture payload\n%%EOF\n');
const png = new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aM1sAAAAASUVORK5CYII=', 'base64'));
const rel = (id, type, target) => `<Relationship Id="${id}" Type="${N.rel}/${type}" Target="${target}"/>`;
const rels = content => `<Relationships xmlns="${N.pkg}">${content}</Relationships>`;
const fill = (id, crop = '') => `<p:blipFill><a:blip r:embed="${id}"/>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill>`;
const picture = content => `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>${content}<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
const slide = content => `<p:sld xmlns:p="${N.p}" xmlns:a="${N.a}" xmlns:r="${N.rel}" xmlns:mc="${N.mc}" xmlns:ma="${ma}" mc:Ignorable="ma"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${content}</p:spTree></p:cSld></p:sld>`;
const open = async files => K.open(new Uint8Array(await (await L.zip.write(files)).arrayBuffer()));
const children = (element, ns, name) => element.children.filter(e => e.namespaceURI === ns && e.localName === name);

async function savePicture(changeCrop = false) {
  const alternate = `<mc:AlternateContent><mc:Choice Requires="ma">${fill('rId1')}</mc:Choice><mc:Fallback>${fill('rId2')}</mc:Fallback></mc:AlternateContent>`;
  const source = await open(Object.entries({
    '[Content_Types].xml': `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="pdf" ContentType="application/pdf"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
    '_rels/.rels': rels(rel('rId1', 'officeDocument', 'ppt/presentation.xml')),
    'ppt/presentation.xml': `<p:presentation xmlns:p="${N.p}" xmlns:r="${N.rel}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels(rel('rId1', 'slide', 'slides/slide1.xml')),
    [part]: slide(picture(alternate)),
    [K.relsPath(part)]: rels(rel('rId1', 'image', '../media/image1.pdf') + rel('rId2', 'image', '../media/image2.png')),
    [pdfPart]: pdf,
    [pngPart]: png,
  }).map(([name, data]) => ({ name, data })));
  const shape = { id: 'picture-1', type: 'image', x: 0, y: 0, w: 100, h: 60, media: pngPart,
    crop: { l: 0, t: 0, r: 0, b: 0 }, img: {}, keep: {} };
  L.properties.shape(shape, source.xml(part).getElementsByTagNameNS(N.p, 'pic')[0], source, part);
  const doc = { slides: [{ shapes: [shape] }], designs: {} };
  L.properties.seal(doc);
  if (changeCrop) shape.crop.l = 0.25;

  const writer = new K.Writer(source, { doc });
  writer.claim(part, 'regenerated');
  writer.carryRels(source, '', '');
  // The ordinary picture generator uses only the PNG. Carrying all slide
  // relationships here would hide whether the alternate fill keeps its PDF.
  const pngRel = source.rels(part).find(r => r.part === pngPart);
  const rid = writer.keepRel(part, { source: source.id, owner: part, ...pngRel });
  const generated = picture(fill(rid, shape.crop.l ? `<a:srcRect l="${Math.round(shape.crop.l * 100000)}"/>` : ''));
  writer.put(part, slide(L.properties.apply(shape, { writer, part }, generated)), source.type(part));
  const saved = await open(writer.finish().files);
  return { saved, doc, pic: saved.xml(part).getElementsByTagNameNS(N.p, 'pic')[0] };
}

function assertImage(saved, fillElement, target, type, bytes) {
  const blips = children(fillElement, N.a, 'blip');
  assert.equal(blips.length, 1);
  const id = blips[0].getAttributeNS(N.rel, 'embed');
  const relationships = saved.rels(part).filter(r => r.id === id);
  assert.equal(relationships.length, 1, `saved blip ${id} resolves to one relationship`);
  const relationship = relationships[0];
  assert.equal(relationship.type, N.rel + '/image');
  assert.equal(relationship.external, false);
  assert.equal(relationship.part, target);
  assert.equal(saved.has(target), true);
  assert.equal(saved.type(target), type);
  assert.deepEqual(saved.bytes(target), bytes);
}

test('an unchanged picture keeps its alternate PDF and PNG fills, relationships and bytes', async () => {
  const { saved, doc, pic } = await savePicture();
  const alternatives = children(pic, N.mc, 'AlternateContent');
  assert.equal(alternatives.length, 1, 'unchanged picture must retain AlternateContent');
  assert.equal(children(pic, N.p, 'blipFill').length, 0, 'no duplicate plain fill');
  assert.deepEqual(pic.children.map(e => e.localName), ['nvPicPr', 'AlternateContent', 'spPr']);
  const alternate = alternatives[0];
  assert.deepEqual(alternate.children.map(e => [e.namespaceURI, e.localName]), [[N.mc, 'Choice'], [N.mc, 'Fallback']]);
  const [choice, fallback] = alternate.children;
  assert.equal(choice.getAttribute('Requires'), 'ma');
  assert.equal(choice.lookupNamespaceURI('ma'), ma);
  const choiceFills = children(choice, N.p, 'blipFill'), fallbackFills = children(fallback, N.p, 'blipFill');
  assert.equal(choiceFills.length, 1);
  assert.equal(fallbackFills.length, 1);
  assertImage(saved, choiceFills[0], pdfPart, 'application/pdf', pdf);
  assertImage(saved, fallbackFills[0], pngPart, 'image/png', png);
  assert.equal(saved.rels(part).length, 2);
  assert.deepEqual(doc.losses, []);
});

test('cropping a picture writes the plain PNG fill and records a silent picture-choice loss', async () => {
  const { saved, doc, pic } = await savePicture(true);
  assert.equal(children(pic, N.mc, 'AlternateContent').length, 0);
  const fills = children(pic, N.p, 'blipFill');
  assert.equal(fills.length, 1, 'crop must use the generated plain fill');
  const crops = children(fills[0], N.a, 'srcRect');
  assert.equal(crops.length, 1);
  assert.equal(crops[0].getAttribute('l'), '25000');
  assertImage(saved, fills[0], pngPart, 'image/png', png);
  assert.equal(saved.rels(part).length, 1);
  assert.equal(saved.has(pdfPart), false);
  const losses = doc.losses.filter(e => e.id.startsWith('picture-choice:'));
  assert.equal(losses.length, 1, 'dropping the alternate version must record picture-choice');
  assert.equal(losses[0].id, 'picture-choice:picture-1');
  assert.equal(losses[0].action, 'drop');
  assert.equal(losses[0].where, part);
  assert.equal(losses[0].phase, 'save');
  assert.equal(losses[0].notify ?? false, false); // notify is optional; omitted means silent.
  assert.deepEqual(K.noticeLosses(doc), []);
});

const alternate = (choice, fallback) => `<mc:AlternateContent><mc:Choice Requires="ma">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
const ignorable = el => {
  const namespaces = new Set();
  for (let node = el; node?.nodeType === 1; node = node.parentNode) {
    for (const prefix of (node.getAttributeNS(N.mc, 'Ignorable') || '').split(/\s+/).filter(Boolean)) namespaces.add(node.lookupNamespaceURI(prefix));
  }
  return [...namespaces].sort();
};
const structure = el => ({
  name: [el.namespaceURI, el.localName],
  // Emitting fragments can repeat the same inherited mc:Ignorable declaration.
  ignorable: ignorable(el),
  attributes: Array.from(el.attributes).filter(a => a.prefix !== 'xmlns' && a.name !== 'xmlns')
    .filter(a => !(a.localName === 'Ignorable' && el.lookupNamespaceURI(a.prefix) === N.mc))
    .map(a => [a.prefix ? el.lookupNamespaceURI(a.prefix) : '', a.localName, a.value]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b))),
  children: el.children.map(structure),
});

async function saveUnchangedShape(xml, type, dependency) {
  const source = await open(Object.entries({
    '[Content_Types].xml': `<Types xmlns="${N.ct}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="inkml" ContentType="application/inkml+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`,
    '_rels/.rels': rels(rel('rId1', 'officeDocument', 'ppt/presentation.xml')),
    'ppt/presentation.xml': `<p:presentation xmlns:p="${N.p}" xmlns:r="${N.rel}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>`,
    'ppt/_rels/presentation.xml.rels': rels(rel('rId1', 'slide', 'slides/slide1.xml')),
    [part]: slide(xml),
    [K.relsPath(part)]: rels(rel('rId1', dependency.type, '../' + dependency.part.slice(4))),
    [dependency.part]: dependency.bytes,
  }).map(([name, data]) => ({ name, data })));
  const tag = type === 'group' ? 'grpSp' : 'sp';
  const original = source.xml(part).getElementsByTagNameNS(N.p, tag)[0];
  const shape = { id: 'object-1', type, x: 0, y: 0, w: 100, h: 60, keep: {}, ...(type === 'group' ? { kids: [] } : {}) };
  L.properties.shape(shape, original, source, part);
  const doc = { slides: [{ shapes: [shape] }], designs: {} };
  L.properties.seal(doc);
  const writer = new K.Writer(source, { doc });
  writer.claim(part, 'regenerated');
  writer.carryRels(source, '', '');
  writer.carryRels(source, part, part);
  // The existing generator's whole-object alternative must stay in its place.
  // The first Phase 4 capture instead tries to insert it as a p:blipFill here.
  let generated;
  assert.doesNotThrow(() => { generated = L.properties.apply(shape, { writer, part }, K.raw(original)); });
  assert.equal(shape.keep.properties.blipFillChoice, undefined, 'whole-object alternatives are not picture-fill properties');
  writer.put(part, slide(generated), source.type(part));
  const saved = await open(writer.finish().files);
  const actual = saved.xml(part).getElementsByTagNameNS(N.p, tag)[0];
  assert.deepEqual(structure(actual), structure(original), 'unedited whole-object XML stays structurally identical');
  assert.equal(children(actual, N.p, 'blipFill').length, 0, 'no fill is inserted into a group or ordinary shape');
  assert.deepEqual(saved.bytes(dependency.part), source.bytes(dependency.part));
  assert.deepEqual(doc.losses, []);
}

test('a stress013-style group containing an alternate whole picture is not a picture fill', async () => {
  const xml = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="Ink group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${alternate('<p:contentPart r:id="rId1"/>', picture('<p:blipFill/>').replace('id="2"', 'id="3"'))}</p:grpSp>`;
  await saveUnchangedShape(xml, 'group', { type: 'ink', part: 'ppt/ink/ink1.inkml',
    bytes: new TextEncoder().encode('<ink xmlns="http://www.w3.org/2003/InkML"><trace>0 0, 1 1</trace></ink>') });
});

test('an ordinary shape with alternate whole spPr and a drawing fill is not a picture fill', async () => {
  const xml = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Pattern shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${alternate('<p:spPr><a:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>', '<p:spPr><a:solidFill><a:srgbClr val="336699"/></a:solidFill></p:spPr>')}</p:sp>`;
  await saveUnchangedShape(xml, 'shape', { type: 'image', part: pngPart, bytes: png });
});
