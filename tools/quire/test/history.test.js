const test = require('node:test');
const assert = require('node:assert/strict');
globalThis.window = globalThis;
require('../../../public/common/core.js');
require('../../../public/quire/js/dmodel.js');
require('../../../public/quire/js/ops.js');
const { D, O } = L;

function fixture() {
  const doc = D.newDoc();
  const textbox = { kind: 'tb', blocks: [D.para([D.text('Text inside the box')])] };
  const shape = D.item('shape', { tb: textbox, w: 120, h: 60 });
  const p = D.para([D.text('Outside'), shape]);
  doc.main.blocks = [p];
  D.doc = doc; D.resetHistory();
  return { doc, p, shape, textbox };
}

test('editing, indexing for save, undo and redo keep textbox contents and owner links', () => {
  const { doc, p, textbox, shape } = fixture();
  D.tx('Type', () => O.insertText(D.pos(p, 7), ' edited'));
  D.reindex(doc); // Layout/save adds the owner link after the before-snapshot was taken.
  assert.equal(textbox.owner, shape);
  assert.equal(D.undo(), true);
  assert.equal(p.runs[0].text, 'Outside');
  assert.equal(p.runs[1].tb.blocks[0].runs[0].text, 'Text inside the box');
  D.reindex(doc);
  assert.equal(p.runs[1].tb.owner, p.runs[1]);
  assert.equal(D.redo(), true);
  assert.equal(p.runs[0].text, 'Outside edited');
  D.reindex(doc);
  assert.equal(p.runs[1].tb.owner, p.runs[1]);
});

test('indexed textbox paragraphs can be copied and whole-document snapshots restored', () => {
  const { doc, p } = fixture();
  D.reindex(doc);
  const snap = D.snapshot(doc);
  const copy = D.cloneBlocks([p])[0];
  assert.notEqual(copy.id, p.id);
  assert.notEqual(copy.runs[1].tb.blocks[0].id, p.runs[1].tb.blocks[0].id);
  assert.equal(copy.runs[1].tb.blocks[0].runs[0].text, 'Text inside the box');
  D.restoreSnapshot(doc, snap);
  D.reindex(doc);
  const shape = doc.main.blocks[0].runs[1];
  assert.equal(shape.tb.owner, shape);
  assert.doesNotThrow(() => JSON.stringify(D.snapshot(doc)));
});
