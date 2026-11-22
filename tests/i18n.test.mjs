// The UI's languages (docs/i18n.md): the English catalog is current, the text that can't be translated yet only
// gets less, and common/i18n.js translates, fills placeholders and falls back to English. The dictionaries here are
// made up for the test; the real translations (public/common/i18n/<lang>.json) are not read.
//   node --test tests/i18n.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { gather } from '../tools/i18n.mjs';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
/* the pieces of English joined together or put in template literals with values, which can't be translated:
   only ever lower this as they are rewritten as one L.t('… {0} …', value) string */
const UNTRANSLATABLE = 6;

test('the English catalog is current: node tools/i18n.mjs gather', () => {
  const { catalog } = gather();
  const saved = JSON.parse(fs.readFileSync(path.join(root, 'public/common/i18n/en.json'), 'utf8'));
  const added = Object.keys(catalog).filter((k) => !(k in saved)), gone = Object.keys(saved).filter((k) => !(k in catalog));
  assert.deepEqual({ added: added.slice(0, 10), gone: gone.slice(0, 10) }, { added: [], gone: [] }, `${added.length} new and ${gone.length} removed strings`);
});

test('text that can\'t be translated yet doesn\'t grow (tmp/i18n/untranslatable.txt)', () => {
  const { fragments } = gather();
  assert.ok(fragments.length <= UNTRANSLATABLE, `${fragments.length} pieces, the limit is ${UNTRANSLATABLE}: write the new text as one L.t('… {0} …', value) string`);
});

test('PDF syntax and font metadata are not UI text; marked export errors are', () => {
  const { catalog } = gather();
  for (const text of ['BT 3 Tr', 'CMapName currentdict /CMap defineresource pop', 'Version 1.0']) {
    assert.equal(Object.hasOwn(catalog, text), false, text);
  }
  for (const text of ['The PDF exporter could not be loaded. Please try again.',
    'This browser blocked the page image.']) {
    assert.ok(Object.hasOwn(catalog, text), text);
  }
  assert.ok(catalog['This browser blocked the page image.'].where.some(where => where.startsWith('common/pdf.js:')));
});

/** common/i18n.js in a page whose UI language is lang, with dictionary dict */
async function page(lang, dict) {
  const html = { getAttribute: () => (lang === 'qps' ? 'qps-ploc' : lang) };
  const context = { window: {}, URL, NodeFilter: { SHOW_TEXT: 4, SHOW_ELEMENT: 1 }, localStorage: { getItem: () => null, setItem() {} },
    document: { documentElement: html, currentScript: { src: 'http://127.0.0.1/common/i18n.js' }, readyState: 'complete', createTreeWalker: () => ({ currentNode: null, nextNode: () => null }) },
    fetch: async (url) => ({ ok: true, json: async () => (String(url).endsWith('/' + (lang === 'qps' ? 'en' : lang) + '.json') ? dict : null) }) };
  context.window.window = context.window;
  vm.runInNewContext(fs.readFileSync(path.join(root, 'public/common/lang.js'), 'utf8'), context);
  vm.runInNewContext(fs.readFileSync(path.join(root, 'public/common/i18n.js'), 'utf8'), context);
  await context.window.VO.i18n.ready;
  return context.window.VO;
}

test('a translation, else the English, with its placeholders filled', async () => {
  const VO = await page('fr', { 'Close': 'Fermer', 'Slide {0} of {1}': 'Diapositive {0} sur {1}', 'Page {0}': 'Page' });
  assert.equal(VO.i18n.lang, 'fr');
  assert.equal(VO.t('Close'), 'Fermer');
  assert.equal(VO.t('Slide {0} of {1}', 2, 9), 'Diapositive 2 sur 9');
  assert.equal(VO.t('Open'), 'Open', 'no translation: the English');
  assert.equal(VO.t('Page {0}', 3), 'Page 3', 'a translation that drops a placeholder is not used');
  assert.equal(VO.t(''), '');
  assert.equal(VO.t(null), null);
});

test('a context after "|" picks a translation of its own, else the plain text\'s', async () => {
  const VO = await page('fr', { 'Line': 'Trait', 'Line|chart': 'Courbes', 'Home': 'Accueil' });
  assert.equal(VO.t('Line|chart'), 'Courbes');
  assert.equal(VO.t('Line'), 'Trait');
  assert.equal(VO.t('Home|key'), 'Accueil', 'no translation of its own: the plain text\'s');
  assert.equal(VO.t('Esc|key'), 'Esc', 'none at all: the English, without its context');
  assert.equal(VO.i18n.plain('Line|chart'), 'Line');
});

test('a "|" that is not a context is text: a file name filled into a message, the user\'s own', async () => {
  for (const [lang, dict] of [['en', null], ['fr', { '\'{0}\' is protected.': '« {0} » est protégé.' }]]) {
    const VO = await page(lang, dict);
    const msg = VO.t('\'{0}\' is protected.', 'A|B.pptx');
    assert.ok(msg.includes('A|B.pptx'), msg);
    assert.equal(VO.t(msg), msg, 'translated again, as L.h does with a message');
    for (const s of ['Revenue | Expenses', 'Q1|Q2', 'a|']) assert.equal(VO.t(s), s);
  }
});

test('English fetches nothing and changes nothing', async () => {
  const VO = await page('en', null);
  assert.equal(VO.t('&File'), '&File');
  assert.equal(VO.t('Slide {0} of {1}', 1, 2), 'Slide 1 of 2');
  assert.equal(VO.t('Line|chart'), 'Line', 'a context is not shown');
});

test('the pseudo-locale changes the catalog\'s strings only, keeping placeholders and access keys', async () => {
  const VO = await page('qps', { '&File': {}, 'Slide {0} of {1}': {} });
  assert.equal(VO.t('&File'), '[&Fílé ··]');
  assert.equal(VO.t('Slide {0} of {1}', 2, 9), '[Slídé 2 óf 9 ···]');
  assert.equal(VO.t('Report.docx'), 'Report.docx', 'text not in the catalog (a file name) stays as it is');
});
