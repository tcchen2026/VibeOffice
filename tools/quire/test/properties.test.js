const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../../public/common/core.js');
require('../../../public/common/xml.js');
require('../../../public/common/opc.js');
require('../../../public/common/opc-order.js');
require('../../../public/quire/js/dmodel.js');
require('../../../public/quire/js/ops.js');
require('../../../public/quire/js/preserve.js');
const { D, O, preserve: P, opc: K } = L;
const w14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const doc = () => { const d = D.newDoc(); D.doc = d; D.resetHistory(); return d; };
const context = d => ({ writer: new K.Writer(null, { doc: d }), doc: d, part: 'word/document.xml' });
const property = (xml, model = {}, kind = 'rPr') => P.properties(model, K.parse(`<w:${kind} xmlns:w="${K.NS.w}" xmlns:w14="${w14}" xmlns:mc="${K.NS.mc}" mc:Ignorable="w14">${xml}</w:${kind}>`), kind, { part: 'word/document.xml' });

test('formatting changes replace their own property and undo restores unsupported effects', () => {
  const d = doc(), r = property('<w:lang w:val="en-US" w:bidi="ar-SA"/><w14:glow w14:rad="190500"/><w14:shadow w14:blurRad="76200"/><w14:reflection w14:blurRad="12700"/>', { lang: 'en-US' });
  const p = D.para([D.text('hello', r)]); d.main.blocks = [p];
  const emit = () => P.propertyXML('rPr', p.runs[0].rPr, '<w:b/><w:lang w:val="en-US"/><w:shadow w:val="0"/>', context(d));
  const original = P.propertyXML('rPr', r, '<w:lang w:val="en-US"/>', context(d));
  D.tx('Change shadow and bold', () => O.setRunProps(D.pos(p, 0), D.pos(p, 5), { shadow: false, b: true }));
  const changed = emit();
  assert.match(changed, /w14:glow/); assert.match(changed, /w14:reflection/);
  assert.doesNotMatch(changed, /w14:shadow/); assert.match(changed, /w:shadow w:val="0"/);
  assert.match(changed, /w:bidi="ar-SA"/);
  assert.ok(changed.indexOf('<w:b ') < changed.indexOf('<w14:glow'));
  D.undo(); assert.equal(P.propertyXML('rPr', p.runs[0].rPr, '<w:lang w:val="en-US"/>', context(d)), original);
  assert.equal(d.losses?.length || 0, 0);
});

test('property alternatives keep both branches and change only the edited property', () => {
  const d = doc(), r = property('<mc:AlternateContent><mc:Choice Requires="w14"><w14:shadow w14:blurRad="76200"/><w14:glow w14:rad="190500"/></mc:Choice><mc:Fallback><w:shadow/></mc:Fallback></mc:AlternateContent>', { shadow: true });
  const before = P.propertyXML('rPr', r, '<w:shadow/>', context(d));
  assert.match(before, /mc:Choice/); assert.match(before, /mc:Fallback/); assert.match(before, /w14:shadow/);
  r.shadow = false;
  const after = P.propertyXML('rPr', r, '<w:b/><w:shadow w:val="0"/>', context(d));
  assert.match(after, /w14:glow/); assert.doesNotMatch(after, /w14:shadow/);
  assert.equal((after.match(/<w:shadow /g) || []).length, 1);
  assert.ok(after.indexOf('<w:b ') < after.indexOf('<mc:AlternateContent'));
  assert.match(after, /mc:Fallback/); assert.equal(d.losses?.length || 0, 0);
});

test('permission copies remap both endpoints, partial copies omit them, and incomplete saves report loss', () => {
  const d = doc(), start = P.rangeMarker(K.parse(`<w:permStart xmlns:w="${K.NS.w}" w:id="42" w:edGrp="everyone" w:colFirst="0" w:colLast="2"/>`), { part: 'word/document.xml' });
  const end = P.rangeMarker(K.parse(`<w:permEnd xmlns:w="${K.NS.w}" w:id="42"/>`), { part: 'word/document.xml' });
  const p = D.para([start, D.text('hello'), end]); d.main.blocks = [p];
  const copy = D.cloneBlocks(O.copyRange(D.pos(p, 0), D.pos(p, 5)))[0];
  d.main.blocks.push(copy);
  const ctx = { ...context(d), ranges: P.prepareRanges(d) };
  const sourceXML = [start, end].map(it => P.markerXML(it, ctx)).join('');
  const copyXML = copy.runs.filter(it => it.t === 'perm').map(it => P.markerXML(it, ctx)).join('');
  const ids = xml => [...xml.matchAll(/w:id="(\d+)"/g)].map(m => m[1]);
  assert.equal(ids(sourceXML)[0], ids(sourceXML)[1]); assert.equal(ids(copyXML)[0], ids(copyXML)[1]);
  assert.notEqual(ids(sourceXML)[0], ids(copyXML)[0]); assert.match(copyXML, /w:colLast="2"/);
  assert.equal(D.cloneBlocks(O.copyRange(D.pos(p, 0), D.pos(p, 2)))[0].runs.some(it => it.t === 'perm'), false);
  p.runs.pop(); ctx.ranges = P.prepareRanges(d);
  assert.equal(P.markerXML(start, ctx), ''); assert.equal(d.losses.length, 1);
});
