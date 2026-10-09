const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../../public/common/core.js');
require('../../../public/common/zip.js');
require('../../../public/common/xml.js');
require('../../../public/common/opc.js');
require('../../../public/quire/js/dmodel.js');
require('../../../public/quire/js/preserve.js');
const { opc: K, preserve: P } = L;
const xml = body => `<w:document xmlns:w="${K.NS.w}">${body}</w:document>`;
const pkg = async body => K.open(await (await L.zip.write([{ name: 'doc.xml', data: xml(body) }])).arrayBuffer());
const save = (source, doc, body) => {
  const p = K.output(source, { doc, audit: P.reportLosses });
  p.bind('doc.xml', 'doc.xml'); p.put('doc.xml', xml(body), 'application/xml');
  assert.equal(K.pendingLosses(doc).length, 0, 'preparing a new save clears earlier writer notices');
  p.finish();
};

test('discarded extra body text has an explicit save notice that survives draft recovery', async () => {
  const first = '<w:body><w:p><w:r><w:t>First body</w:t></w:r></w:p></w:body>';
  const second = '<w:body><w:p><w:r><w:t>Additional body</w:t></w:r></w:p></w:body>';
  const source = await pkg(first + second), doc = {};
  save(source, doc, first + second); assert.equal(doc.losses.length, 0);
  save(source, doc, first); assert.equal(doc.losses.length, 1);
  assert.match(doc.losses[0].what, /outside the main document and won't be saved.*\(1\)/);
  const recovered = {}; K.recoverLosses(recovered, K.lossState(doc));
  assert.equal(K.noticeLosses(recovered)[0].what, doc.losses[0].what, 'the user is still told after draft recovery');
  save(source, doc, first + second); assert.equal(doc.losses.length, 0, 'a preserving save clears the cancelled conversion');
});

test('only omitted body-level final sections produce the malformed-structure notice', async () => {
  const ordinarySection = '<w:p><w:pPr><w:sectPr/></w:pPr></w:p>';
  const source = await pkg('<w:body>' + ordinarySection + '<w:sectPr/><w:sectPr/></w:body>'), doc = {};
  save(source, doc, '<w:body>' + ordinarySection + '<w:sectPr/></w:body>');
  assert.equal(doc.losses.length, 1);
  assert.match(doc.losses[0].what, /Conflicting final section definitions.*\(1\)/);
  assert.equal(K.noticeLosses(doc).length, 0, 'a repaired structure is recorded, not shown');
  const valid = await pkg('<w:body>' + ordinarySection + '<w:sectPr/></w:body>');
  save(valid, doc, '<w:body><w:sectPr/></w:body>');
  assert.equal(doc.losses.length, 0, 'ordinary section deletion is not an extra-final-section loss');
});
