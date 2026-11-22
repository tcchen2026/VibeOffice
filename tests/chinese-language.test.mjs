// Chinese UI selection, including query/storage precedence and dictionary URLs. No real translations are read.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const read = async (p) => {
  if (!process.env.BASE) return fs.readFileSync(new URL('../public/' + p, import.meta.url), 'utf8');
  const r = await fetch(process.env.BASE + '/' + p);
  assert.ok(r.ok, p);
  return r.text();
};
const pages = ['index.html', 'lite.html', 'quire/index.html', 'ledger/index.html', 'lectern/index.html'];
async function boot(page, { query = '', stored = 'auto', languages = ['en-US'], language = 'en-US', unavailable = false } = {}) {
  const html = await read(page), attrs = { lang: 'en', dataset: {}, getAttribute(k) { return this[k]; } };
  const fetched = [], state = new Map([['vibeoffice.lang', stored]]);
  const c = { URL, URLSearchParams, location: { search: query }, navigator: { languages, language },
    localStorage: { getItem: (k) => { if (unavailable) throw Error('unavailable'); return state.get(k); }, setItem: (k, v) => state.set(k, v) },
    document: { documentElement: attrs, readyState: 'loading', addEventListener() {}, currentScript: { src: 'http://127.0.0.1/common/i18n.js' } },
    fetch: async (url) => { fetched.push(String(url)); return { ok: true, json: async () => ({ Close: 'TEST' }) }; },
    addEventListener() {},
  };
  c.window = c;
  vm.createContext(c);
  const head = html.slice(0, html.indexOf('common/looks.css'));
  for (const m of head.matchAll(/<script(?:\s+src="([^"]+)")?>([\s\S]*?)<\/script>/g)) {
    vm.runInContext(m[1] ? await read(m[1].replace(/^\.\.\//, '')) : m[2], c);
  }
  vm.runInContext(await read('common/i18n.js'), c);
  await c.VO.i18n.ready;
  return { c, attrs, fetched, state };
}

test('every page: Chinese regions and scripts select the canonical Chinese language', async () => {
  const cases = [
    ['zh', 'zh-CN'], ['zh-CN', 'zh-CN'], ['zh-SG', 'zh-CN'], ['zh-Hans', 'zh-CN'], ['zh-Hans-SG', 'zh-CN'], ['zh-Hans-TW', 'zh-CN'], ['zh-MY', 'zh-CN'],
    ['zh-TW', 'zh-Hant'], ['ZH-tw', 'zh-Hant'], ['zh-HK', 'zh-Hant'], ['zh-MO', 'zh-Hant'],
    ['zh-Hant', 'zh-Hant'], ['ZH-hAnt', 'zh-Hant'], ['zh-Hant-TW', 'zh-Hant'], ['zh-Hant-HK', 'zh-Hant'],
  ];
  for (const page of pages) for (const [tag, expected] of cases) {
    const { c, attrs, fetched } = await boot(page, { languages: [tag, 'en-US'] });
    assert.equal(attrs.lang, expected, page + ': ' + tag);
    assert.equal(c.VO.i18n.lang, expected);
    assert.equal(c.VO.i18n.dates, expected);
    assert.deepEqual(fetched, ['http://127.0.0.1/common/i18n/' + expected + '.json']);
    assert.equal(c.VO.t('Close'), 'TEST');
  }
});

test('query overrides stored language, storage overrides browser, and Auto follows supported preferences', async () => {
  for (const page of pages) {
    assert.equal((await boot(page, { query: '?lang=zh-TW', stored: 'zh-CN', languages: ['en-US'] })).attrs.lang, 'zh-Hant');
    assert.equal((await boot(page, { query: '?lang=zh-HK', stored: 'zh-CN' })).attrs.lang, 'zh-Hant');
    assert.equal((await boot(page, { stored: 'zh-TW', languages: ['zh-CN'] })).attrs.lang, 'zh-Hant');
    assert.equal((await boot(page, { stored: 'auto', languages: ['es', 'zh-TW', 'en'] })).attrs.lang, 'zh-Hant');
    assert.equal((await boot(page, { query: '?lang=zh-TW', unavailable: true })).attrs.lang, 'zh-Hant');
    assert.equal((await boot(page, { languages: ['fr-FR', 'zh-TW'] })).attrs.lang, 'fr');
    assert.equal((await boot(page, { languages: null, language: 'zh-HK' })).attrs.lang, 'zh-Hant');
    assert.equal((await boot(page, { query: '?lang=auto', stored: 'zh-CN', languages: ['zh-MO'] })).attrs.lang, 'zh-Hant');
    const { c, state } = await boot(page);
    assert.ok(c.VO.lang.list().some((l) => l.id === 'zh-CN'));
    assert.ok(c.VO.lang.list().some((l) => l.id === 'zh-Hant'));
    c.VO.lang.set('zh-Hant');
    assert.equal(state.get('vibeoffice.lang'), 'zh-Hant');
    c.VO.lang.set('zh-CN');
    c.VO.lang.set('zh-HK');
    assert.equal(state.get('vibeoffice.lang'), 'zh-Hant', 'aliases are stored with the canonical ID');
    const legacy = await boot(page, { stored: 'zh-TW' });
    legacy.c.VO.lang.set('zh-Hant');
    assert.equal(legacy.state.get('vibeoffice.lang'), 'zh-Hant');
    c.VO.lang.set('auto');
    assert.equal(c.VO.lang.get(), 'auto');
  }
});
