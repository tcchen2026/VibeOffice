// East Asian fonts (docs/i18n.md, Fonts): the names Japanese, Korean and Chinese Office write are the fonts every table
// knows by their English names, a run draws its East Asian characters from its East Asian font, and Ledger sizes rows
// for a Japanese or Korean Excel's default font:  node --test tests/fonts.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import fs from 'node:fs/promises';
import vm from 'node:vm';

const require = createRequire(import.meta.url);
globalThis.window = globalThis;
require('../public/common/core.js');
require('../public/ledger/js/model.js');
const L = globalThis.L;

test('localized names are the English fonts', () => {
  assert.equal(L.fontKey('ＭＳ 明朝'), 'ms mincho');
  assert.equal(L.fontKey('맑은 고딕'), 'malgun gothic');
  assert.equal(L.fontKey('宋体'), 'simsun');
  assert.equal(L.fontKey('Times New Roman'), 'times new roman');
  assert.ok(L.fontIsEA('游ゴシック') && L.fontIsEA('Batang') && !L.fontIsEA('Century'));
});

test('Simplified Chinese open-source font references retain their families and East Asian fallbacks', async () => {
  const source = process.env.BASE ? await (await fetch(process.env.BASE + '/common/core.js')).text()
    : await fs.readFile(new URL('../public/common/core.js', import.meta.url), 'utf8');
  const refs = [
    ['文泉驿微米黑', 'WenQuanYi Micro Hei', 'sans-serif'],
    ['思源黑体', 'Source Han Sans SC', 'sans-serif'],
    ['思源宋体', 'Source Han Serif SC', 'serif'],
    ['Noto Sans CJK SC', 'Noto Sans CJK SC', 'sans-serif'], ['Noto Sans SC', 'Noto Sans SC', 'sans-serif'],
    ['Noto Serif CJK SC', 'Noto Serif CJK SC', 'serif'], ['Noto Serif SC', 'Noto Serif SC', 'serif'],
  ];
  for (const lang of ['zh-CN', 'en', 'zh-Hant']) {
    const c = { document: { documentElement: { getAttribute: () => lang } } }; c.window = c;
    vm.runInNewContext(source, c);
    for (const [name, family, generic] of refs) {
      assert.equal(c.L.FONT_LIST.includes(name), lang === 'zh-CN', name + ': ' + lang);
      assert.equal(c.L.fontKey(name), family.toLowerCase());
      assert.ok(c.L.fontIsEA(name) && c.L.fontIsEA(family), name);
      const stack = c.L.fontStack('Times New Roman', name);
      assert.ok(stack.includes('"' + name + '"') && stack.includes('"' + family + '"'), stack);
      assert.ok(stack.endsWith(',' + generic), stack);
      const serif = generic === 'serif', counterpart = serif ? '"Noto Serif CJK SC"' : '"Noto Sans CJK SC"';
      if (!family.startsWith('WenQuanYi')) assert.ok(stack.indexOf(counterpart) < stack.indexOf(serif ? '"SimSun"' : '"Microsoft YaHei"'), stack);      assert.ok(c.L.fontStack(name).startsWith('"' + name + '",'), 'the chosen family is tried first');
    }
    // Resolve an unavailable family with a controlled set of installed fonts, independently of this machine.
    c.L.fontInstalled = (f) => ['Noto Sans CJK SC', 'Noto Serif CJK SC', 'Microsoft YaHei', 'SimSun'].includes(f);
    assert.equal(c.L.fontUsed('思源黑体'), 'Noto Sans CJK SC');
    assert.equal(c.L.fontUsed('思源宋体'), 'Noto Serif CJK SC');
    assert.equal(c.L.fontUsed('文泉驿微米黑'), 'Microsoft YaHei');
    c.L.fontInstalled = (f) => ['WenQuanYi Micro Hei', 'Source Han Sans SC', 'Source Han Serif SC'].includes(f);
    assert.equal(c.L.fontUsed('文泉驿微米黑'), 'WenQuanYi Micro Hei');
    assert.equal(c.L.fontUsed('思源宋体'), 'Source Han Serif SC');
    assert.equal(c.L.fontUsed('宋体'), 'Source Han Serif SC');
    c.L.fontInstalled = () => false;
    assert.equal(c.L.fontUsed('思源黑体'), 'sans-serif');
    assert.equal(c.L.fontUsed('思源宋体'), 'serif');
  }
});

test('a run\'s letters come from its Latin font, its East Asian characters from its East Asian font', () => {
  const stack = L.fontStack('Century', 'ＭＳ 明朝');
  assert.ok(stack.startsWith('"Century",'), stack);
  assert.ok(stack.indexOf('"ＭＳ 明朝"') > stack.indexOf('"Century"'), 'the East Asian font after the Latin one');
  assert.ok(stack.includes('"MS Mincho"') && stack.endsWith(',serif'), stack);
  assert.equal(stack.match(/,serif/g).length, 1, 'one generic family, at the end');
  assert.equal(L.fontStack('Arial', 'Arial'), L.fontStack('Arial'), 'the same font once');
  assert.equal(L.fontStack('Arial', 'Wingdings'), L.fontStack('Arial'), 'not a font for East Asian text');
});

test('East Asian text is told from Latin text', () => {
  for (const s of ['日本語', 'ひらがな', 'カタカナ', '한국어', '中文', '（全角）']) assert.ok(L.hasEAText(s), s);
  for (const s of ['English', 'Français', 'Größe', '', null]) assert.ok(!L.hasEAText(s), String(s));
});

test('Ledger: the rows of a Japanese or Korean Excel\'s default font', () => {
  assert.equal(L.model.fontRowPt({ name: 'ＭＳ Ｐゴシック', sz: 11 }), 13.5);
  assert.equal(L.model.fontRowPt({ name: '游ゴシック', sz: 11 }), 18.75);
  assert.equal(L.model.fontRowPt({ name: '맑은 고딕', sz: 11 }), 16.5);
  assert.equal(L.model.fontRowPt({ name: 'Calibri', sz: 11 }), 15, 'Latin fonts as before');
});

test('Wingdings and Symbol characters show as Office draws them, also when Word stores them in the private-use area', () => {
  // I24 E1: a converted Word 97-2003 list's bullets, Wingdings U+F076 and Symbol U+F0A8 (a dot and a missing glyph before)
  assert.equal(L.mapSymbolChar('', 'Wingdings'), '❖');
  assert.equal(L.mapSymbolChar('v', 'Wingdings'), '❖');
  assert.equal(L.mapSymbolChar('', 'Wingdings'), '➢');
  assert.equal(L.mapSymbolChar('', 'Symbol'), '♦');
  assert.equal(L.mapSymbolChar('', 'Symbol'), '•');
  assert.equal(L.mapSymbolChar('a', 'Symbol'), 'α');
  assert.equal(L.mapSymbolChar('', 'Symbol'), '⇒');
  assert.equal(L.mapSymbolChar('•', 'Symbol'), '•', 'a Unicode character stays as it is');
  assert.equal(L.mapSymbolChar('ð', 'Symbol'), 'ð', 'Symbol has nothing at 0xF0');
});
