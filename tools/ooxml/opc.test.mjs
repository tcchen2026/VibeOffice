import test from 'node:test';
import assert from 'node:assert/strict';
import '../../public/common/xml.js';
import '../../public/common/opc.js';
import '../../public/common/opc-order.js';
import '../../public/common/zip.js';
import '../../public/common/dml.js';

const K = L.opc, N = K.NS;
const rel = (id, type, target, external = false) => `<Relationship Id="${id}" Type="${type}" Target="${K.esc(target)}"${external ? ' TargetMode="External"' : ''}/>`;
const rels = content => `<Relationships xmlns="${N.pkg}">${content}</Relationships>`;
const types = '<Types xmlns="' + N.ct + '"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>';
async function pkg(parts) {
  const blob = await L.zip.write(Object.entries({ '[Content_Types].xml': types, ...parts }).map(([name, data]) => ({ name, data })));
  return K.open(new Uint8Array(await blob.arrayBuffer()));
}
const text = data => typeof data === 'string' ? data : new TextDecoder().decode(data);
const kid = (el, name) => el.children.find(c => c.localName === name);
const children = xml => K.parse(xml).children.map(e => e.localName);

test('baseline bytes are immutable; the graph retains full types, sharing, cycles and external URLs', async () => {
  const p = await pkg({ 'a.xml': '<a/>', 'b.xml': '<b/>', 'shared.bin': new Uint8Array([3, 4]),
    '_rels/.rels': rels(rel('rId40', N.rel + '/officeDocument', 'a.xml')),
    '_rels/a.xml.rels': rels(rel('rId1', 'urn:vendor/type', 'b.xml') + rel('alias', N.rel + '/image', 'shared.bin') + rel('web', N.rel + '/hyperlink', 'https://example.invalid/a?b=1&c=2', true)),
    '_rels/b.xml.rels': rels(rel('back', 'urn:vendor/back', 'a.xml') + rel('same', 'urn:vendor/bytes', 'shared.bin')) });
  assert.equal(p.main, 'a.xml'); assert.equal(p.rels('a.xml')[0].type, 'urn:vendor/type');
  assert.deepEqual(new Set(K.subtree(p, 'a.xml').parts), new Set(['a.xml', 'b.xml', 'shared.bin']));
  const bytes = p.bytes('shared.bin'); bytes[0] = 99; assert.equal(p.bytes('shared.bin')[0], 3);
  assert.throws(() => { p.rels('a.xml')[0].target = 'wrong.xml'; }, TypeError);
  const w = new K.Writer(p); w.carryRels(p, '', ''); const result = w.finish();
  for (const name of ['a.xml', 'b.xml', 'shared.bin', '_rels/a.xml.rels', '_rels/b.xml.rels']) assert.deepEqual(new Uint8Array(result.files.find(f => f.name === name).data), p.bytes(name));
  assert.match(w.rels('').xml(), /Id="rId40"/); assert.equal(result.dropped.length, 0);
});

test('opaque backlinks follow a renamed regenerated part; unfilled claims fail explicitly', async () => {
  const p = await pkg({ 'a.xml': '<a/>', 'b.xml': '<b/>', '_rels/.rels': rels(rel('rId9', N.rel + '/officeDocument', 'a.xml')),
    '_rels/a.xml.rels': rels(rel('rId2', 'urn:child', 'b.xml')), '_rels/b.xml.rels': rels(rel('rId7', 'urn:back', 'a.xml')) });
  const w = new K.Writer(p); w.claim('a.xml', 'regenerated', 'new.xml'); w.carryRels(p, '', ''); w.carryRels(p, 'a.xml', 'new.xml');
  assert.throws(() => w.finish(), /Unwritten relationship target/);
  w.put('new.xml', '<a edited="1"/>', 'application/xml'); const files = w.finish().files;
  assert.match(text(files.find(f => f.name === '_rels/b.xml.rels').data), /Id="rId7"[^>]+Target="new.xml"/);
  assert.equal(text(files.find(f => f.name === 'b.xml').data), '<b/>');
  assert.throws(() => w.claim('b.xml', 'regenerated', 'new.xml'), /collision/);
});

test('copying an opaque notes page redirects its slide backlink and shares immutable targets', async () => {
  const p = await pkg({ 'slide.xml': '<slide/>', 'notes.xml': '<notes/>', 'master.xml': '<master/>',
    '_rels/notes.xml.rels': rels(rel('slide', N.rel + '/slide', 'slide.xml') + rel('master', N.rel + '/notesMaster', 'master.xml')) });
  const w = new K.Writer(p); w.claim('slide.xml'); w.put('slide.xml', '<slide/>', 'application/xml'); w.put('copy.xml', '<slide/>', 'application/xml');
  w.carry(p, 'notes.xml'); w.copyPart(p, 'notes.xml', 'copied-notes.xml', { 'slide.xml': 'copy.xml' });
  assert.deepEqual(w.parts.get('copied-notes.xml'), p.bytes('notes.xml'));
  assert.match(text(w.parts.get('_rels/copied-notes.xml.rels')), /Id="slide"[^>]+Target="copy.xml"/);
  assert.match(text(w.parts.get('_rels/copied-notes.xml.rels')), /Id="master"[^>]+Target="master.xml"/);
  assert.doesNotThrow(() => w.finish());
});

test('failed carry leaves no partial output and respects regenerated dependency boundaries', async () => {
  const p = await pkg({ 'good.xml': '<good/>', 'broken.xml': '<broken/>',
    '_rels/broken.xml.rels': rels(rel('one', 'urn:child', 'good.xml') + rel('two', 'urn:child', 'absent.xml')) });
  const w = new K.Writer(p), names = [...w.names], mapping = [...w.mapping];
  assert.throws(() => w.carry(p, 'broken.xml'), /Missing dependency: absent.xml/);
  assert.equal(w.parts.size, 0); assert.equal(w.carried.size, 0);
  assert.deepEqual([...w.names], names); assert.deepEqual([...w.mapping], mapping);
  w.claim('broken.xml'); w.put('broken.xml', '<converted/>', 'application/xml');
  assert.equal(w.carry(p, 'broken.xml'), 'broken.xml');
  assert.equal(w.parts.get('broken.xml'), '<converted/>');
  assert.doesNotThrow(() => w.finish());
});

test('rId allocation reserves original gaps and preserves aliases for one shared target', () => {
  const r = new K.Rels([{ id: 'rId83', type: 'urn:a', target: 'a.xml' }, { id: 'rId2', type: 'urn:a', target: 'a.xml' }]);
  assert.equal(r.add('urn:a', 'a.xml'), 'rId83');
  assert.equal(r.add('urn:a', 'a.xml', false, 'rId2'), 'rId2');
  assert.equal(r.add('urn:other', 'b.xml'), 'rId84');
  assert.throws(() => r.add('urn:wrong', 'x.xml', false, 'rId2'), /rebound/);
});

test('output aliases preserve opaque relationship IDs without overwriting their graphs', async () => {
  const p = await pkg({ 'original/main.xml': '<main/>', 'original/theme.xml': '<theme/>', 'original/data.xml': '<data/>',
    '_rels/.rels': rels(rel('main40', N.rel + '/officeDocument', 'original/main.xml')),
    'original/_rels/main.xml.rels': rels(rel('theme7', N.rel + '/theme', '/original/theme.xml')),
    'original/_rels/theme.xml.rels': rels(rel('original-id', 'urn:data', 'data.xml')) });
  const out = K.output(p, { consumes: (base, r) => r.type === N.rel + '/officeDocument' });
  out.bind('generated/main.xml', p.main);
  out.bind('generated/theme.xml', 'original/theme.xml', 'opaque');
  out.rels('').add(N.rel + '/officeDocument', 'generated/main.xml');
  assert.equal(out.rels('generated/main.xml').add(N.rel + '/theme', 'theme.xml'), 'theme7');
  // A native writer can have prepared a different rels file before ownership
  // decides to carry the opaque target. It must not replace the carried graph.
  out.rels('generated/theme.xml').add('urn:wrong', 'absent.xml');
  out.put('generated/theme.xml', '<wrong/>');
  out.put('generated/main.xml', '<main edited="1"/>', 'application/xml');
  const result = out.finish();
  assert.equal(result.dropped.length, 0);
  assert.equal(out.writer.rels('original/main.xml').list.length, 1);
  assert.deepEqual(out.writer.parts.get('original/theme.xml'), p.bytes('original/theme.xml'));
  assert.deepEqual(out.writer.parts.get('original/_rels/theme.xml.rels'), p.bytes('original/_rels/theme.xml.rels'));
  assert.equal(out.writer.parts.has('generated/main.xml'), false);
});

test('draft recovery retains unacknowledged conversions after the source parts are gone', () => {
  const original = {}, w = new K.Writer(null, { doc: original });
  w.loss({ id: 'converted:1', what: 'Converted object', where: 'Object 1', action: 'conversion' });
  K.loss(original, { id: 'edited:2', what: 'Edited object', action: 'drop' });
  K.acknowledge(original, original.losses.slice(1));
  const recovered = {};
  K.recoverLosses(recovered, JSON.parse(JSON.stringify(K.lossState(original))));
  new K.Writer(null, { doc: recovered });
  assert.equal(recovered.losses.length, 2);
  assert.deepEqual(K.pendingLosses(recovered).map(e => e.id), ['converted:1']);
});

test('rewritten roots retain their original entity prolog and Strict namespace form', () => {
  const original = '<?xml version="1.0"?><!DOCTYPE root [<!ENTITY name "A &amp; B">]><?before keep?><root>&name;</root><?after keep?>';
  const output = K.partXML(original, '<root changed="1">&name;</root>');
  assert.equal(output, original.replace('<root>', '<root changed="1">'));
  const strict = K.strictXML(`<w:document xmlns:w="${N.w}" xmlns:r="${N.rel}"><w:body note="${N.w}"/></w:document>`);
  assert.match(strict, /xmlns:w="http:\/\/purl.oclc.org\/ooxml\/wordprocessingml\/main"/);
  assert.match(strict, /xmlns:r="http:\/\/purl.oclc.org\/ooxml\/officeDocument\/relationships"/);
  assert.ok(strict.includes(`note="${N.w}"`)); // ordinary values are not URI declarations
});

test('identity scopes survive transfers, explicit index maps and bounded slide/master ranges', () => {
  const ids = new K.Identities(); ids.reserve('slide2.xml', 'shape', '99');
  assert.equal(ids.resolve('pkg', 'slide1.xml', 'shape', '2', { primary: 'pkg', scope: 'slide2.xml' }), '100');
  assert.equal(ids.resolve('pkg', 'slide1.xml', 'shape', '2', { primary: 'pkg', scope: 'slide2.xml' }), '100');
  ids.bind('pkg', 'styles', 'dxf', '3', '7');
  assert.equal(ids.resolve('pkg', 'styles', 'dxf', '3', { primary: 'pkg' }), '7');
  assert.equal(ids.fresh('presentation', 'sldId'), '256');
  assert.equal(ids.fresh('presentation', 'sldMasterId'), '2147483648');
  ids.reserve('presentation', 'sldId', '2147483647');
  assert.equal(ids.fresh('presentation', 'sldId'), '257');
  ids.reserve('lexical', 'shape', '+001');
  assert.equal(ids.fresh('lexical', 'shape'), '2');
  assert.equal(ids.key('pkg', 'lexical', 'shape', '01'), ids.key('pkg', 'lexical', 'shape', '1'));
});

test('fragments retain lexical XML, default namespaces, QName values and inner prefix rebinding', () => {
  const source = `<root xmlns="urn:outer" xmlns:mc="${N.mc}" xmlns:q="urn:outer-q" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" mc:Ignorable="q" xml:space="preserve"><node><!--keep--><q:data xmlns:q='urn:inner' xsi:type='q:Type' q:unusual='yes'> <![CDATA[a < b]]> </q:data></node></root>`;
  const el = kid(K.parse(source), 'node'), raw = K.raw(el), normalized = K.serialize(el);
  assert.match(raw, /<!--keep-->/); assert.match(raw, /<!\[CDATA\[a < b\]\]>/); assert.match(raw, /xmlns:q='urn:inner'/);
  assert.match(raw, /xml:space="preserve"/); assert.match(raw, /mc:Ignorable="q"/);
  for (const value of [raw, normalized]) {
    const node = K.parse(value); assert.equal(node.namespaceURI, 'urn:outer'); assert.equal(node.children[0].namespaceURI, 'urn:inner');
    assert.equal(node.children[0].lookupNamespaceURI('q'), 'urn:inner');
  }
});

test('AlternateContent is captured before inner wrappers collapse, once per outer record', () => {
  const xml = `<root xmlns:mc="${N.mc}" xmlns:wps="urn:wps" xmlns:new="urn:new"><mc:AlternateContent><mc:Choice Requires="wps"><one/><mc:AlternateContent><mc:Choice Requires="new"><new:thing/></mc:Choice><mc:Fallback><two/></mc:Fallback></mc:AlternateContent></mc:Choice><mc:Fallback><fallback/></mc:Fallback></mc:AlternateContent><mc:AlternateContent><mc:Fallback/></mc:AlternateContent></root>`;
  const tree = K.parse(xml); L.dml.resolveAC(tree.parentNode, xml);
  const a = K.alternate(tree.children[0]), b = K.alternate(tree.children[1]);
  assert.equal(a.id, b.id); assert.match(a.xml, /new:thing/); assert.match(a.xml, /<fallback\/>/);
  assert.equal(K.alternates(tree).length, 2); // includes the empty alternate for the reader's anchor
  assert.deepEqual(tree.children.map(c => c.localName), ['one', 'two']);
});

test('relationships and all definitions/referrers share one identity map through duplicate and undo JSON', async () => {
  const part = 'ppt/slides/slide1.xml';
  const xml = `<p:sld xmlns:p="${N.p}" xmlns:a="${N.a}" xmlns:r="${N.rel}"><p:cNvPr id="100"/><p:grpSp><p:cNvPr id="8"/><p:pic><p:cNvPr id="9"/><a:blip r:embed="rId20" r:link="url"/></p:pic><p:cTn id="15"/><p:tn val="15"/><p:spTgt spid="9"/><a:stCxn id="100"/></p:grpSp></p:sld>`;
  const p = await pkg({ [part]: xml, 'ppt/media/video.bin': new Uint8Array([1, 2, 3]), [K.relsPath(part)]: rels(rel('rId20', 'urn:media/video', '../media/video.bin') + rel('url', N.rel + '/hyperlink', 'https://example.invalid/', true)) });
  const original = K.fragment(kid(K.parse(xml), 'grpSp'), { pkg: p, part });
  const fragment = JSON.parse(JSON.stringify(original)), w = new K.Writer(p);
  w.claim(part); w.put(part, xml, 'application/xml');
  const unchanged = w.emit(fragment, part); assert.match(unchanged, /id="8"/); assert.match(unchanged, /r:embed="rId20"/);
  const copy = K.duplicate({ fragment }).fragment, changed = K.parse(w.emit(copy, part));
  const get = name => changed.getElementsByTagName('*').filter(e => e.localName === name);
  const shapeIds = get('cNvPr').map(e => e.getAttribute('id'));
  assert.deepEqual(shapeIds, ['101', '102']);
  assert.equal(get('spTgt')[0].getAttribute('spid'), '102');
  assert.equal(get('stCxn')[0].getAttribute('id'), '100'); // external connector target was not duplicated
  assert.equal(get('cTn')[0].getAttribute('id'), get('tn')[0].getAttribute('val'));
  assert.notEqual(get('cTn')[0].getAttribute('id'), '15');
  assert.deepEqual(w.parts.get('ppt/media/video.bin'), new Uint8Array([1, 2, 3]));
  const again = w.emit(JSON.parse(JSON.stringify(copy)), part); assert.equal(again, K.serialize(changed));
});

test('multiple models from one alternate emit the original choice and fallback exactly once', async () => {
  const xml = `<root xmlns:mc="${N.mc}" xmlns:new="urn:new"><mc:AlternateContent><mc:Choice Requires="new"><new:media/></mc:Choice><mc:Fallback><a/><b/></mc:Fallback></mc:AlternateContent></root>`;
  const p = await pkg({ 'content.xml': xml }), tree = K.parse(xml); L.dml.resolveAC(tree.parentNode, xml);
  const w = new K.Writer(p), fragments = tree.children.map(e => K.fragment(e, { pkg: p, part: 'content.xml' }));
  assert.match(w.emit(fragments[0], 'content.xml'), /new:media/);
  assert.equal(w.emit(fragments[1], 'content.xml'), '');
  const copies = K.duplicate(fragments);
  assert.match(w.emit(copies[0], 'content.xml'), /new:media/); assert.equal(w.emit(copies[1], 'content.xml'), '');
});

test('foreign dependency graphs are transported as bytes and renamed without collisions', async () => {
  const p = await pkg({ 'media1.bin': new Uint8Array([1]) });
  const source = await pkg({ 'media1.bin': new Uint8Array([2]), 'metadata.xml': '<x/>', '_rels/metadata.xml.rels': rels(rel('original', 'urn:binary', 'media1.bin')) });
  const bundle = JSON.parse(JSON.stringify(K.export([{ source: source.id, part: 'metadata.xml' }]))), imported = K.import(bundle)[source.id];
  const w = new K.Writer(p); w.carry(p, 'media1.bin'); w.carry(imported, 'metadata.xml');
  const result = w.finish(); assert.equal(result.dropped.length, 0);
  assert.deepEqual(w.parts.get('media1.bin'), new Uint8Array([1])); assert.deepEqual(w.parts.get('media2.bin'), new Uint8Array([2]));
  assert.match(text(w.parts.get('_rels/metadata.xml.rels')), /Id="original"[^>]+Target="media2.bin"/);
  assert.equal(w.types.get('media2.bin'), 'application/octet-stream');
});

test('foreign fragment references are rebased with their bytes and dangling copied targets are reported', async () => {
  const xml = `<p:pic xmlns:p="${N.p}" xmlns:a="${N.a}" xmlns:r="${N.rel}"><p:cNvPr id="7"/><a:blip r:embed="rId1"/></p:pic>`;
  const source = await pkg({ 'shape.xml': xml, 'image.bin': new Uint8Array([8]), '_rels/shape.xml.rels': rels(rel('rId1', N.rel + '/image', 'image.bin')) });
  const f = K.fragment(K.parse(xml), { pkg: source, part: 'shape.xml' });
  const mapping = K.import(K.export([{ source: source.id, part: 'image.bin' }]));
  const rebased = K.remapSources(K.duplicate(f), mapping), dest = await pkg({ 'main.xml': '<x/>' }), w = new K.Writer(dest);
  const out = w.emit(rebased, 'main.xml'); assert.match(out, /r:embed="rId1"/); assert.deepEqual(w.parts.get('image.bin'), new Uint8Array([8]));
  const dangling = K.fragment(K.parse(`<p:cxnSp xmlns:p="${N.p}" xmlns:a="${N.a}"><p:cNvPr id="3"/><a:stCxn id="7"/></p:cxnSp>`), { pkg: source, part: 'shape.xml' });
  assert.throws(() => w.emit(K.duplicate(dangling), 'main.xml'), /Unresolved copied shape/);
  assert.equal(K.pendingLosses(w.doc).length, 1);
});

test('the loss ledger acknowledges each entry independently and reports invalidated signatures', async () => {
  const p = await pkg({ 'main.xml': '<main/>', 'signature.xml': '<signature/>', '_rels/.rels': rels(rel('doc', N.rel + '/officeDocument', 'main.xml') + rel('sig', N.pkg + '/digital-signature/origin', 'signature.xml')) });
  const doc = {}, w = new K.Writer(p, { doc }); w.carryRels(p, '', ''); w.finish();
  assert.ok(doc.losses.some(e => e.id.startsWith('signature:'))); assert.ok(!w.parts.has('signature.xml'));
  K.acknowledge(doc, K.pendingLosses(doc)); assert.equal(K.pendingLosses(doc).length, 0);
  K.loss(doc, { id: 'new', what: 'Edited opaque content', where: 'slide 1', action: 'conversion' });
  assert.equal(K.pendingLosses(doc).length, 1);
  K.acknowledge(doc, K.pendingLosses(doc)); K.loss(doc, { id: 'new', what: 'Deleted opaque content', where: 'slide 1', action: 'drop' });
  assert.equal(K.pendingLosses(doc).length, 1);
  K.attach(doc, p); assert.equal(JSON.parse(JSON.stringify(doc)).pkg, undefined); assert.equal(doc.pkg, p);
});

test('save losses are recomputed without clearing reader or edit losses and their acknowledgements', () => {
  const doc = {};
  K.loss(doc, { id: 'edited', what: 'The user replaced an unsupported object.', action: 'conversion' });
  const first = new K.Writer(null, { doc });
  first.loss({ id: 'macro', what: 'The selected format does not support macros.', action: 'drop' });
  K.acknowledge(doc, K.pendingLosses(doc));
  new K.Writer(null, { doc });
  assert.deepEqual(doc.losses.map(e => e.id), ['edited']);
  assert.deepEqual(K.pendingLosses(doc), []);
  const next = new K.Writer(null, { doc });
  next.loss({ id: 'new', what: 'A new object cannot be saved.', action: 'drop' });
  assert.deepEqual(K.pendingLosses(doc).map(e => e.id), ['new']);
});

test('property bags keep lexical unknown values, types and IDs while owned entries change', () => {
  const xml = '<Properties xmlns="urn:p"><property name="int" pid="7"><i4>12</i4></property><!--retained--><property name="text" pid="14"><s>old</s></property><unknown flag=\'yes\'/></Properties>';
  const out = K.mergeBag(xml, { text: '<property name="text" pid="14"><s>new</s></property>', added: '<property name="added" pid="15"><b>true</b></property>' }, e => e.getAttribute('name'));
  assert.equal(out, xml.replace('<s>old</s>', '<s>new</s>').replace('</Properties>', '<property name="added" pid="15"><b>true</b></property></Properties>'));
  assert.equal(K.mergeBag('<Properties xmlns="urn:p"/>', { x: '<child/>' }), '<Properties xmlns="urn:p"><child/></Properties>');
  assert.equal(K.relationshipType(N.strictRel + '/extendedProperties'), N.rel + '/extended-properties');
  assert.equal(K.relationshipType(N.strictRel + '/customProperties'), N.rel + '/custom-properties');
  assert.equal(K.relationshipType('urn:vendor/customProperties'), 'urn:vendor/customProperties');
  const a = K.attributes('<settings foo="foo" other=\'untouched\'><unknown/></settings>', { foo: null, added: '1' });
  assert.equal(a, '<settings  other=\'untouched\' added="1"><unknown/></settings>');
  assert.equal(K.attributes('<settings a=\'old\'/>', { a: 'new' }), '<settings a=\'new\'/>');
});

test('schema merge replaces just the changed property and keeps extensions at their original anchor', () => {
  const untouched = `<w:settings xmlns:w="${N.w}"><w:zoom/><w:view/><w:vendor/></w:settings>`;
  assert.equal(K.merge(untouched, {}, 'w:CT_Settings'), untouched);
  const xml = `<a:spPr xmlns:a="${N.a}" xmlns:v="urn:vendor"><a:xfrm/><v:unknown flag="yes"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><!--kept--><a:effectLst><a:glow rad="12700"/></a:effectLst><a:sp3d/></a:spPr>`;
  const out = K.merge(xml, { ['{' + N.a + '}solidFill']: [], ['{' + N.a + '}noFill']: '<a:noFill/>' }, 'a:CT_ShapeProperties');
  assert.deepEqual(children(out), ['xfrm', 'unknown', 'noFill', 'effectLst', 'sp3d']);
  assert.match(out, /<!--kept-->/); assert.match(out, /glow rad="12700"/); assert.match(out, /flag="yes"/);
  const effect = `<a:effectLst xmlns:a="${N.a}"><a:glow rad="3"/><a:outerShdw blurRad="4"/><a:reflection blurRad="5"/></a:effectLst>`;
  const replaced = K.merge(effect, { ['{' + N.a + '}outerShdw']: '<a:outerShdw blurRad="99"/>' }, 'a:CT_EffectList');
  assert.match(replaced, /glow rad="3"/); assert.match(replaced, /reflection blurRad="5"/); assert.match(replaced, /blurRad="99"/);
});

test('geometry edits change the outer group transform without rewriting child geometry or effects', () => {
  const xml = `<p:grpSp xmlns:p="${N.p}" xmlns:a="${N.a}"><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="254000"/><a:chOff x="33" y="44"/><a:chExt cx="55" cy="66"/></a:xfrm></p:grpSpPr><p:sp><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:glow rad="70"/></p:spPr></p:sp></p:grpSp>`;
  const changed = K.setBox(xml, { x: 3, y: 4, w: 20, h: 40 });
  assert.equal(changed, xml.replace('x="0" y="0"', 'x="38100" y="50800"').replace('cx="127000" cy="254000"', 'cx="254000" cy="508000"'));
  assert.equal(K.getBox(changed).x, 3); assert.equal(K.getBox(changed).w, 20);
  assert.throws(() => K.setBox(xml, { w: NaN }), /Invalid box/);
});

test('editing a Strict property replaces it once and retains Strict namespace and sibling effects', () => {
  const ns = 'http://purl.oclc.org/ooxml/drawingml/main';
  const xml = `<a:effectLst xmlns:a="${ns}"><a:glow rad="3"/><a:outerShdw blurRad="4"/><a:reflection blurRad="5"/></a:effectLst>`;
  const out = K.merge(xml, { ['{' + N.a + '}outerShdw']: `<a:outerShdw xmlns:a="${N.a}" blurRad="99"/>` }, 'a:CT_EffectList');
  const nodes = K.parse(out).children;
  assert.deepEqual(nodes.map(e => e.localName), ['glow', 'outerShdw', 'reflection']);
  assert.ok(nodes.every(e => e.namespaceURI === ns));
  assert.equal(nodes[1].getAttribute('blurRad'), '99');
  assert.match(out, /glow rad="3"/); assert.match(out, /reflection blurRad="5"/);
});

test('all Office variants retain their distinct main content type and macro-enabled MIME', async () => {
  assert.equal(Object.keys(K.formats).length, 14);
  assert.equal(new Set(Object.values(K.formats).map(f => f.contentType)).size, 14);
  for (const [extension, info] of Object.entries(K.formats)) {
    const p = await pkg({ 'main.xml': '<main/>', '_rels/.rels': rels(rel('main', N.rel + '/officeDocument', 'main.xml')),
      '[Content_Types].xml': types.replace('</Types>', `<Override PartName="/main.xml" ContentType="${info.contentType}"/></Types>`) });
    assert.equal(K.variant(p), extension);
    assert.equal(K.format({ pkg: p }, null, 'docx').ext, extension);
    assert.equal(info.macro, extension.endsWith('m'));
    assert.equal(info.mime.includes('macroEnabled.12'), info.macro);
    assert.equal(K.format({ pkg: p }, 'xlsx', 'docx').ext, 'xlsx');
  }
  assert.equal(K.formats.dotm.contentType, 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml');
});

test('Word anchor position and extent update while wrapping and picture offsets survive', () => {
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const xml = `<wp:anchor xmlns:wp="${wp}" xmlns:a="${N.a}"><wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>12700</wp:posOffset></wp:positionV><wp:extent cx="25400" cy="38100"/><wp:wrapSquare wrapText="bothSides"/><a:xfrm><a:off x="0" y="0"/><a:ext cx="25400" cy="38100"/></a:xfrm></wp:anchor>`;
  const out = K.setBox(xml, { x: 9, y: 10, w: 11, h: 12 });
  assert.match(out, /<wp:posOffset>114300<\/wp:posOffset>/); assert.match(out, /<wp:posOffset>127000<\/wp:posOffset>/);
  assert.match(out, /wrapText="bothSides"/); assert.match(out, /a:off x="0" y="0"/);
  assert.equal(K.getBox(out).h, 12);
});

test('spreadsheet anchors and VML fractional cell anchors update without changing payloads', () => {
  const ns = K.KNOWN_NS.xdr;
  const xml = `<xdr:twoCellAnchor xmlns:xdr="${ns}"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:clientData fLocksWithSheet="1"/></xdr:twoCellAnchor>`;
  const out = K.setBox(xml, { from: { r: 5, cOff: 2 } }); assert.equal(K.getBox(out).from.r, 5); assert.equal(K.getBox(out).from.cOff, 2); assert.match(out, /fLocksWithSheet="1"/);
  const vml = `<v:shape xmlns:v="${K.KNOWN_NS.v}" xmlns:x="${K.KNOWN_NS.x}" style="position:absolute;margin-left:2pt;top:3pt;width:4pt;height:5pt;z-index:7"><x:ClientData ObjectType="Button"><x:Anchor>0,0,1,0,2,0,3,0</x:Anchor></x:ClientData></v:shape>`;
  const moved = K.setBox(vml, { x: 10, w: 6, from: { c: 2, r: 5, cFraction: 0.5, rFraction: 0.25 } });
  assert.match(moved, /margin-left:10pt/); assert.match(moved, /width:6pt/); assert.match(moved, /z-index:7/); assert.match(moved, /2, 512, 5, 64, 2, 0, 3, 0/);
});

test('part URI resolution handles relative, absolute and escaped names without escaping the package', () => {
  assert.deepEqual(K.resolve('ppt/slides/slide1.xml', '../media/a%20b.bin#sample'), { part: 'ppt/media/a b.bin', fragment: '#sample' });
  assert.equal(K.relative('ppt/slides/slide1.xml', 'ppt/media/a b.bin'), '../media/a%20b.bin');
  assert.deepEqual(K.resolve('ppt/slides/slide1.xml', '/customXml/item.xml'), { part: 'customXml/item.xml', fragment: '' });
  assert.throws(() => K.resolve('', '../../outside'), /escapes/);
  assert.throws(() => K.resolve('a.xml', 'https://example.invalid/'), /Invalid internal/);
});

test('escaped ZIP entry names and relationship URIs share one logical part identity', async () => {
  const p = await pkg({ 'a.xml': '<a/>', 'media/Cort%C3%A1zar%20one.bin': new Uint8Array([7, 8]),
    '_rels/.rels': rels(rel('rId1', N.rel + '/officeDocument', 'a.xml')),
    '_rels/a.xml.rels': rels(rel('rId2', N.rel + '/audio', 'media/Cort%C3%A1zar%20one.bin')) });
  assert.ok(p.has('media/Cortázar one.bin'));
  const w = new K.Writer(p); w.carryRels(p, '', ''); const result = w.finish();
  assert.ok(result.files.some(f => f.name === 'media/Cort%C3%A1zar%20one.bin'));
  const saved = await K.open(new Uint8Array(await (await L.zip.write(result.files)).arrayBuffer()));
  assert.deepEqual(saved.bytes('media/Cortázar one.bin'), new Uint8Array([7, 8]));
  assert.equal(saved.rels('a.xml')[0].part, 'media/Cortázar one.bin');
});
