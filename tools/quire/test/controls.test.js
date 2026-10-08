const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../../public/common/core.js');
require('../../../public/common/xml.js');
require('../../../public/common/opc.js');
require('../../../public/quire/js/dmodel.js');
require('../../../public/quire/js/ops.js');
require('../../../public/quire/js/preserve.js');
const { D, O, preserve: P, opc: K } = L;

function fixture(properties = '<w:text/>', prefix = '', suffix = '') {
  const doc = D.newDoc(); doc.keep = {};
  const control = P.control(K.parse(`<w:sdt xmlns:w="${K.NS.w}"><w:sdtPr><w:id w:val="42"/>${properties}</w:sdtPr><w:sdtContent/></w:sdt>`), { part: 'word/document.xml' }, doc);
  const p = D.para([D.text(prefix), D.item('sdts', { control }), D.text('hello'), D.item('sdte', { key: control.key }), D.text(suffix)]);
  D.normalize(p); doc.main.blocks = [p]; D.doc = doc; D.resetHistory();
  return { doc, p, control };
}
const xml = (doc, record) => P.controlXML(record.control, '<w:r><w:t>value</w:t></w:r>', { writer: new K.Writer(null, { doc }), doc, part: 'word/document.xml' });

test('editing a bound value, undo and redo include the XML-store update', () => {
  const { doc, p, control } = fixture();
  control.binding = { source: 'source', part: 'customXml/item1.xml', storeID: 'store', xpath: '/root/value', mappings: '' };
  D.tx('Type', () => O.insertText(D.pos(p, 2), '!'));
  assert.equal(doc.keep.boundUpdates['store:/root/value'].value, 'he!llo');
  assert.equal(P.controls(doc).records.get(control.key).complete, true);
  D.undo(); assert.equal(doc.keep.boundUpdates, undefined); assert.equal(D.ptext(p), 'hello');
  D.redo(); assert.equal(doc.keep.boundUpdates['store:/root/value'].value, 'he!llo');
});

test('content and deletion locks roll back changes without a history entry', () => {
  for (const [lock, operation] of [
    ['contentLocked', p => O.insertText(D.pos(p, 2), '!')],
    ['sdtLocked', p => O.deleteRange(D.pos(p, 0), D.pos(p, 5))],
    ['sdtContentLocked', p => O.splitPara(D.pos(p, 2))],
  ]) {
    const { doc, p } = fixture(`<w:lock w:val="${lock}"/><w:text/>`);
    D.reindex(doc);
    const before = D.snapshot(doc);
    assert.equal(D.tx('Locked edit', () => operation(p)), false);
    D.reindex(doc); assert.deepEqual(D.snapshot(doc), before); assert.equal(D.canUndo(), false);
  }
});

test('partial deletion unwraps the control and undo restores its exact properties', () => {
  const { doc, p, control } = fixture('<w:alias w:val="Customer"/><w:text/>', 'before ', ' after');
  const original = xml(doc, P.controls(doc).records.get(control.key));
  D.tx('Delete', () => O.deleteRange(D.pos(p, 9), D.pos(p, 15)));
  assert.equal(P.controls(doc).records.size, 0); assert.equal(doc.losses.length, 1);
  D.undo(); assert.equal(xml(doc, P.controls(doc).records.get(control.key)), original);
  assert.equal(doc.losses.length, 0);
});

test('splitting a full-paragraph control promotes it to balanced block boundaries; join and undo keep it', () => {
  const { doc, p, control } = fixture('<w:richText/>');
  D.tx('Split', () => O.splitPara(D.pos(p, 2)));
  let record = P.controls(doc).records.get(control.key);
  assert.equal(record.complete, true); assert.equal(record.text, 'he\nllo');
  assert.equal(p.pPr.sdts[0].key, control.key); assert.equal(doc.losses, undefined);
  D.tx('Join', () => O.joinNext(p));
  record = P.controls(doc).records.get(control.key);
  assert.equal(record.complete, true); assert.equal(record.text, 'hello');
  D.undo(); assert.equal(doc.main.blocks.length, 2);
  D.undo(); assert.equal(p.runs[0].t, 'sdts'); assert.equal(doc.main.blocks.length, 1);
});

test('copying whole controls gives fresh identities; partial copies have no dangling markers', () => {
  const { doc, p, control } = fixture();
  const copy = D.cloneBlocks(O.copyRange(D.pos(p, 0), D.pos(p, 5)))[0];
  const c = copy.runs.find(it => it.t === 'sdts').control;
  assert.notEqual(c.key, control.key);
  const writer = new K.Writer(null, { doc });
  const original = P.controlXML(control, '', { writer, doc, part: 'word/document.xml' });
  const copied = P.controlXML(c, '', { writer, doc, part: 'word/document.xml' });
  assert.notEqual(original.match(/w:val="(\d+)"/)[1], copied.match(/w:val="(\d+)"/)[1]);
  for (const [a, b] of [[0, 2], [2, 5]]) {
    const partial = D.cloneBlocks(O.copyRange(D.pos(p, a), D.pos(p, b)))[0];
    assert.equal(partial.runs.some(it => it.t === 'sdts' || it.t === 'sdte'), false);
  }
  p.runs.push(D.item('be', { id: 'outside' }));
  assert.equal(D.cloneBlocks([p])[0].runs.some(it => it.t === 'be'), false);
  control.binding = { source: 'other-document', part: 'customXml/item1.xml', storeID: '{SOURCE}' };
  const first = D.cloneBlocks([p])[0].runs.find(it => it.t === 'sdts').control.binding;
  const second = D.cloneBlocks([p])[0].runs.find(it => it.t === 'sdts').control.binding;
  assert.notEqual(first.storeID, control.binding.storeID);
  assert.equal(first.storeID, second.storeID, 'repeated copies sharing a store must agree on its new identity');
});

test('typed values change through the value command; typing converts and undo restores the type', () => {
  const { doc, p, control } = fixture('<w:dropDownList><w:listItem w:value="a" w:displayText="Alpha"/><w:listItem w:value="b" w:displayText="Beta"/></w:dropDownList>');
  control.binding = { source: 'source', part: 'customXml/item1.xml', storeID: 'store', xpath: '/root/value', mappings: '' };
  D.tx('Choose', () => P.setControlValue(control.key, 'b'));
  assert.equal(D.ptext(p), 'Beta'); assert.equal(doc.keep.boundUpdates['store:/root/value'].value, 'b');
  assert.equal(control.converted, undefined); assert.equal(doc.losses, undefined);
  D.tx('Type', () => O.insertText(D.pos(p, 2), '!'));
  let record = P.controls(doc).records.get(control.key);
  assert.equal(record.control.converted, true); assert.equal(record.control.binding, undefined);
  assert.equal(xml(doc, record).includes('dropDownList'), false);
  D.undo(); record = P.controls(doc).records.get(control.key);
  assert.equal(record.control.converted, undefined); assert.equal(record.control.binding.storeID, 'store');
});
