// Run the existing password test bodies on Intern's CDP/server, without modifying
// app tests or closing any pre-existing Chromium context or tab.
const fs = require('node:fs'), path = require('node:path'), Module = require('node:module');
const { chromium } = require('playwright');
const mode = process.argv[2], BASE = 'http://127.0.0.1:8769', CDP = 'http://127.0.0.1:9223';
const output = path.resolve('tmp/legacy/artifacts/password-ui'); fs.mkdirSync(output, { recursive: true });
let remote, owned = [];
const prepare = async page => {
  const goto = page.goto.bind(page); page.goto = (url, options) => goto(url.replace('http://127.0.0.1:8760', BASE), options);
  await page.route(BASE + '/__legacy_fixture__/**', async route => {
    const name = decodeURIComponent(new URL(route.request().url()).pathname.slice('/__legacy_fixture__/'.length));
    const files = {
      'tika__5479add191f2__testPPT_protected_passtika.pptx': path.join(output, 'testPPT_protected_passtika.pptx'),
      'oxsdk__e255bb9bc6d4__typical.pptx': '/work/corpora/powerpoint/openxml-sdk__88d7f7c81442__typical.pptx',
    };
    if (!files[name]) return route.abort();
    await route.fulfill({ status: 200, contentType: 'application/octet-stream', headers: { 'access-control-allow-origin': BASE }, body: fs.readFileSync(files[name]) });
  });
  return page;
};
const facade = {
  async newContext(options) { const context = await remote.newContext(options); owned.push(context);
    const page = context.newPage.bind(context); context.newPage = async () => prepare(await page()); return context; },
  async newPage(options) { const context = await facade.newContext(options); return context.newPage(); },
  async close() { for (const context of owned) await context.close(); owned = []; await remote.close(); },
};
const adapted = { chromium: { async launch() { remote = await chromium.connectOverCDP(CDP); return facade; } } };
let file, source;
if (mode === 'ledger') {
  file = path.resolve('tests/ledger/ui/password.js'); source = fs.readFileSync(file, 'utf8');
  const fixture = path.join(output, 'protected_passtika.xlsx'); fs.copyFileSync('/work/corpora/excel/poi__5764ea6be01c__protected_passtika.xlsx', fixture);
  process.argv = [process.execPath, file, fixture, 'tika'];
} else if (mode === 'lectern') {
  file = path.resolve('tests/lectern/ui-features.js'); source = fs.readFileSync(file, 'utf8');
  // Only its password sections are requested; later chart fixtures belong to a retired corpus.
  source = source.slice(0, source.indexOf('  /* 3. imported chart')) + `
  ok('no page errors', errors.length === 0, errors);
  fs.writeFileSync(path.join(out, 'checks.json'), JSON.stringify(checks, null, 2));
  await b.close();
  if (checks.some(c => !c.ok)) process.exitCode = 1;
})().catch(async e => { console.error(e); await globalThis.__i24Close?.(); process.exitCode = 1; });`;
  process.argv = [process.execPath, file, path.join(output, 'lectern')];
  process.env.HOST = BASE + '/lectern/'; process.env.CORPUS_URL = BASE + '/__legacy_fixture__/';
} else throw new Error('Usage: node tools/legacy/password-ui.cjs ledger|lectern');
globalThis.__i24Close = () => remote ? facade.close() : undefined;
process.on('unhandledRejection', async e => { console.error(e); await globalThis.__i24Close(); process.exitCode = 1; });
const mod = new Module(file, module); mod.filename = file; mod.paths = Module._nodeModulePaths(path.dirname(file));
const normal = mod.require.bind(mod); mod.require = name => name === 'playwright' ? adapted : normal(name); mod._compile(source, file);
