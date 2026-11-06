// A font this computer doesn't have and the suite has no stand-in for is drawn with what the document's font table
// says, as Word does: its alternate name (w:altName), else a face of its kind (a serif one for w:family="roman", a
// fixed-pitch one for w:pitch="fixed"). The SDK's complex0.docx names "Bremen Bd BT" (altName Courier New): Word 2016
// prints it in Courier New, Quire drew it in Arial; POI's bug59058.docx now shows Word's 52 pages (53 before).
// node --test tests/font-stand-ins.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
require('../public/common/core.js');
const L = globalThis.L;

test('a missing font takes the stand-in its document names, else one of its kind', () => {
  assert.match(L.fontStack('Bremen Bd BT'), /^"Bremen Bd BT",Arial,/, 'no hint: Arial, as before');
  L.fontHints.set(L.fontKey('Bremen Bd BT'), { alt: 'Courier New', family: 'decorative', pitch: 'variable' });
  L.fontHints.set(L.fontKey('CG Times (W1)'), { alt: null, family: 'roman', pitch: 'variable' });
  L.fontHints.set(L.fontKey('Letter Gothic'), { alt: null, family: 'swiss', pitch: 'fixed' });
  L.fontHints.set(L.fontKey('Algerian'), { alt: null, family: 'decorative', pitch: 'variable' });
  assert.match(L.fontStack('Bremen Bd BT'), /^"Bremen Bd BT","Courier New","Cousine",/);
  assert.match(L.fontStack('CG Times (W1)'), /^"CG Times \(W1\)","Times New Roman","Tinos",/);
  assert.match(L.fontStack('Letter Gothic'), /^"Letter Gothic","Courier New",/);
  assert.match(L.fontStack('Algerian'), /^"Algerian",Arial,/, 'no stand-in of its kind: Arial');
  // our own stand-ins come first
  L.fontHints.set('calibri', { alt: 'Arial', family: 'swiss' });
  assert.match(L.fontStack('Calibri'), /^"Calibri","Carlito",/);
});
