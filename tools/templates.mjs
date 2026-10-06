// Build the Start Center's template gallery from the apps themselves: for each app, its template list
// (VO.templates(), from the app's own catalog) and a preview of Blank and of every template, drawn by
// the app (VO.thumbnail()). Writes public/common/templates.json and public/<app>/templates/<id>.jpg.
// Run it after adding, removing or changing a template.
//
//   node tools/templates.mjs
//
// Needs the Chromium at CDP 127.0.0.1:9222 (see docs/testing.md); starts tools/serve.py if needed.
// Env: CDP (default http://127.0.0.1:9222), BASE (default http://127.0.0.1:8760)
import fs from 'fs';
import path from 'path';
import net from 'net';
import { spawn } from 'child_process';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const CDP = process.env.CDP || 'http://127.0.0.1:9222';
const BASE = process.env.BASE || 'http://127.0.0.1:8760';
const APPS = { quire: 'Blank Document', ledger: 'Blank Workbook', lectern: 'Blank Presentation' };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const port = +new URL(BASE).port;
const listening = () => new Promise((r) => { const s = net.connect(port, '127.0.0.1', () => { s.end(); r(true); }).on('error', () => r(false)); });
let server = null;
if (!(await listening())) {
  server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port)], { stdio: 'ignore' });
  for (let i = 0; i < 50 && !(await listening()); i++) await sleep(100);
}

const target = await (await fetch(`${CDP}/json/new?about:blank`, { method: 'PUT' })).json();
const ws = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((r, j) => { ws.onopen = r; ws.onerror = j; });
let seq = 0;
const pending = new Map(), errors = [];
ws.onmessage = (e) => {
  const m = JSON.parse(e.data);
  if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); return; }
  if (m.method === 'Runtime.exceptionThrown') errors.push(m.params.exceptionDetails.exception?.description ?? m.params.exceptionDetails.text);
};
const send = (method, params = {}) => new Promise((res) => { const id = ++seq; pending.set(id, res); ws.send(JSON.stringify({ id, method, params })); });
const evaluate = async (expression) => {
  const r = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
  if (r.result?.exceptionDetails) throw new Error(r.result.exceptionDetails.exception?.description ?? r.result.exceptionDetails.text);
  return r.result?.result?.value;
};
/** load an app page, wait until it has attached to the suite and drawn */
async function open(url) {
  await send('Page.navigate', { url: `${BASE}/${url}` });
  for (let i = 0; i < 80; i++) {
    await sleep(250);
    if (await evaluate(`document.readyState === 'complete' && !!window.VO && VO.templates().length > 0`).catch(() => false)) break;
  }
  await sleep(1200);
}
const JPEG = `(async () => { const b = await VO.thumbnail(); if (!b) return null; const a = new Uint8Array(await b.arrayBuffer()); let s = ''; for (const x of a) s += String.fromCharCode(x); return btoa(s); })()`;

let failed = false;
try {
  for (const d of ['Page', 'Runtime']) await send(`${d}.enable`);
  await send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 800, deviceScaleFactor: 1, mobile: false });
  const out = {};
  for (const [app, blankName] of Object.entries(APPS)) {
    const dir = path.join(root, 'public', app, 'templates');
    fs.rmSync(dir, { recursive: true, force: true });
    fs.mkdirSync(dir, { recursive: true });
    await open(`${app}/?new`);
    const list = [{ id: '', name: blankName }, ...(await evaluate('VO.templates()'))];
    out[app] = [];
    for (const t of list) {
      if (t.id) await open(`${app}/?template=${encodeURIComponent(t.id)}`);
      const b64 = await evaluate(JPEG);
      const file = `${t.id || 'blank'}.jpg`;
      if (b64) fs.writeFileSync(path.join(dir, file), Buffer.from(b64, 'base64'));
      else { console.log(`${app}/${t.id}: no preview`); failed = true; }
      out[app].push({ id: t.id, name: t.name, img: b64 ? `${app}/templates/${file}` : null });
      console.log(`${app.padEnd(8)} ${(t.id || '(blank)').padEnd(10)} ${t.name}`);
    }
  }
  fs.writeFileSync(path.join(root, 'public/common/templates.json'), JSON.stringify(out, null, 1) + '\n');
  console.log('wrote public/common/templates.json');
} catch (e) {
  failed = true;
  console.log('ABORTED:', e.message);
} finally {
  for (const e of errors) console.log('[exception]', e.slice(0, 400));
  await fetch(`${CDP}/json/close/${target.id}`).catch(() => {});
  ws.close();
  server?.kill();
}
if (failed || errors.length) process.exit(1);
