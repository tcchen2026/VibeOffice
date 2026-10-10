// The look settings (public/common/looks.css) stay complete and the chrome uses them:
//   node --test tools/looks.test.mjs
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const root = new URL('../public/', import.meta.url);
const read = (p) => fs.readFileSync(new URL(p, root), 'utf8');
const looks = read('common/looks.css');
const styleOf = (html) => [...html.matchAll(/<style>([\s\S]*?)<\/style>/g)].map((m) => m[1]).join('\n');
const SHEETS = {
  'common/ui.css': read('common/ui.css'),
  'ledger/index.html': styleOf(read('ledger/index.html')),
  'quire/index.html': styleOf(read('quire/index.html')),
  'lectern/index.html': styleOf(read('lectern/index.html')),
  'index.html': styleOf(read('index.html')),
};
/* set per element by the apps for document content, not by a look */
const DOCUMENT_VARS = new Set(['--tabc', '--olvl', '--bu-c', '--bu-f', '--bu-s', '--bu-w', '--icon-swatch']);

const names = (css) => new Set([...css.matchAll(/(--[a-z0-9-]+)\s*:/g)].map((m) => m[1]));
/** look id → the settings its blocks set (a block's selector may list several looks) */
const settingsOf = () => {
  const out = new Map();
  for (const m of looks.matchAll(/^([^\n{}]*data-look[^\n{}]*)\{([^}]*)\}/gm)) {
    if (!/^\s*(:root|\[data-look)/.test(m[1]) || /\)\s+[.#a-z]/.test(m[1])) continue;   // value blocks, not shape rules
    for (const id of m[1].matchAll(/data-look="([a-z0-9]+)"/g)) out.set(id[1], new Set([...(out.get(id[1]) || []), ...names(m[2])]));
  }
  return out;
};

test('every look sets every setting', () => {
  const all = settingsOf();
  const classic = all.get('classic');
  assert.ok(classic && classic.size > 50, 'Classic block found');
  const ids = [...read('common/suite.js').matchAll(/\{ id: '([a-z0-9]+)', name:/g)].map((m) => m[1]);
  assert.ok(ids.length >= 3, 'looks listed in suite.js');
  for (const id of ids) {
    assert.ok(all.has(id), `look "${id}" has settings in looks.css`);
    const missing = [...classic].filter((n) => !all.get(id).has(n));
    assert.deepEqual(missing, [], `look "${id}" falls back to Classic for ${missing.join(', ')}`);
  }
});

test('every setting the chrome uses is defined', () => {
  const defined = names(looks);
  for (const [file, css] of Object.entries(SHEETS)) {
    const used = new Set([...css.matchAll(/var\((--[a-z0-9-]+)/g)].map((m) => m[1]));
    const missing = [...used].filter((n) => !defined.has(n) && !DOCUMENT_VARS.has(n));
    assert.deepEqual(missing, [], `${file} uses undefined settings`);
  }
});

test('the shared chrome has no colours of its own', () => {
  /* colour swatches and proofing marks show content colours, not the look */
  const allowed = /^\.(cp-sw|cc-|cpk-sw|sg-bad|sg-gram)/;
  const bad = SHEETS['common/ui.css'].split('\n')
    .filter((l) => /#[0-9a-fA-F]{3,8}\b|rgba?\(/.test(l) && !allowed.test(l.trim()));
  assert.deepEqual(bad, []);
});

test('the pages pick the look before the first paint and load looks.css first', () => {
  for (const page of ['index.html', 'quire/index.html', 'ledger/index.html', 'lectern/index.html']) {
    const html = read(page);
    const script = html.indexOf("localStorage.getItem('vibeoffice.look')");
    const sheet = html.indexOf('common/looks.css');
    assert.ok(script > 0 && script < sheet, `${page}: the look is picked before looks.css loads`);
    assert.ok(sheet > 0 && (html.indexOf('common/ui.css') < 0 || sheet < html.indexOf('common/ui.css')), `${page}: looks.css before ui.css`);
  }
});
