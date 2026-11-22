// The lite edition's example documents (desktop/lite-examples/; docs/desktop.md, Lite edition): each app's sample,
// saved by the app itself as the file it would write, which a new ~/VibeOffice/Documents starts with
// (desktop/src/lite.rs, EXAMPLES). Run it after changing a sample, then rebuild the lite app.
//
//   node tools/lite-examples.mjs
//
// Needs the Chromium at CDP 127.0.0.1:9222 (see docs/testing.md); starts tools/serve.py if needed.
import fs from 'node:fs';
import path from 'node:path';
import { ensureServer, openPage } from './ooxml/cdp.mjs';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const out = path.join(root, 'desktop/lite-examples');
const BASE = process.env.BASE || 'http://127.0.0.1:8760';
/** app -> the file its sample becomes */
const EXAMPLES = { quire: 'Welcome to Quire.docx', ledger: 'Quarterly Sales.xlsx', lectern: 'Harbor & Pine.pptx' };
const pause = (ms) => new Promise((r) => setTimeout(r, ms));

const stop = await ensureServer(BASE);
fs.mkdirSync(out, { recursive: true });
try {
  for (const [app, name] of Object.entries(EXAMPLES)) {
    const t = await openPage(app, { persistence: false });
    try {
      // the plain page shows the app's sample (common/suite.js, VO.launch)
      await t.evaluate('delete window.__corpusHooks; true');
      await t.send('Page.navigate', { url: `${BASE}/${app}/?lang=en` });
      for (let i = 0; i < 200 && !(await t.evaluate('!!window.__corpusHooks').catch(() => false)); i++) await pause(100);
      await pause(1500);
      const b64 = await t.evaluate(`(async () => {
        const blob = await __corpusHooks.snapshot(__corpusHooks.current());
        const bytes = new Uint8Array(blob instanceof Blob ? await blob.arrayBuffer() : blob);
        let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
        return btoa(s);
      })()`);
      const bytes = Buffer.from(b64, 'base64');
      if (bytes.length < 1000 || bytes.readUInt32LE(0) !== 0x04034b50) throw new Error(`${app}: not a package`);
      fs.writeFileSync(path.join(out, name), bytes);
      console.log(`${name}: ${bytes.length} bytes`);
      if (t.errors.length) throw new Error(`${app}: ${t.errors.join('; ')}`);
    } finally { await t.close(); }
  }
} finally { stop(); }
