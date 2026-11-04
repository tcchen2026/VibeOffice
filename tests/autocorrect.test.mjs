// AutoCorrect's default list (public/common/spell.js, shared by the three apps): Apache OpenOffice's English entries
// and our own rules, nothing that clashes with an AutoFormat option:  node --test tests/autocorrect.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
globalThis.window = globalThis;
require('../public/common/core.js');
require('../public/common/spell.js');
const list = globalThis.L.spell.autocorrectDefaults();

test('typos from Apache OpenOffice\'s list', () => {
  for (const [k, v] of [['teh', 'the'], ['recieve', 'receive'], ['accomodate', 'accommodate'], ['definately', 'definitely'], ['alot', 'a lot'],
    ['could of been', 'could have been'], ['cutsomers', 'customers'], ['i', 'I']]) assert.equal(list.get(k), v, k);
  assert.ok(list.size > 850, `${list.size} entries`);
});

test('our own: symbols, contractions without their apostrophe, I, accents in loanwords', () => {
  for (const [k, v] of [['(c)', '©'], ['(r)', '®'], ['(tm)', '™'], ['...', '…'], [':)', '☺'], ['-->', '→'], ['==>', '⇒'], ['<=>', '⇔'],
    ['dont', "don't"], ['doesnt', "doesn't"], ['wouldnt', "wouldn't"], ['youre', "you're"], ['theyve', "they've"], ['thats', "that's"],
    ["i'm", "I'm"], ["i've", "I've"], ['cafe', 'café'], ['naive', 'naïve'], ['facade', 'façade'], ['deja vu', 'déjà vu']]) assert.equal(list.get(k), v, k);
});

test('nothing that is a word of its own, and nothing an AutoFormat option does', () => {
  for (const k of ['cant', 'wont', 'well', 'were', 'hell', 'shell', 'shed', 'wed', 'its', 'lets', 'ill', 'id', 'resume', '--', '1/2', '1/4', '3/4'])
    assert.equal(list.has(k), false, k);
  for (const [k, v] of list) assert.ok(k && v && !k.includes('|') && !v.includes('|'), `${k}=${v}`);
});

test('the user\'s entries over the defaults less an app\'s: added, changed, deleted; a deleted own entry leaves no trace', () => {
  const kept = new Map();
  globalThis.VO = { settings: { get: (k) => kept.get(k) ?? null, set: (k, v) => { if (v == null) kept.delete(k); else kept.set(k, v); return true; }, remove: (k) => kept.delete(k) } };
  try {
    const SP = globalThis.L.spell;
    SP.setAutocorrectEntry('zzqx', 'zed queue');
    SP.setAutocorrectEntry('teh', 'tea');
    SP.setAutocorrectEntry('recieve', null);
    assert.equal(SP.autocorrectList().get('zzqx'), 'zed queue');
    SP.setAutocorrectEntry('zzqx', null);
    const l = SP.autocorrectList(['i']);
    assert.equal(l.get('teh'), 'tea');
    assert.equal(l.has('recieve'), false);
    assert.equal(l.has('zzqx'), false);
    assert.equal(l.has('i'), false);
    assert.deepEqual(JSON.parse([...kept.values()][0]), { add: [['teh', 'tea']], del: ['recieve'] });
  } finally { delete globalThis.VO; }
});
