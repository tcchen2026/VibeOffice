// Drive an app page in the Chromium at CDP 127.0.0.1:9222: run JS steps, take screenshots,
// and report console errors, exceptions and any request that leaves 127.0.0.1.
//
//   node tools/shot.mjs <page> [steps.json | 'inline json'] [--w 1280 --h 800]
//                       [--qs query] [--ready 'expr'] [--out dir] [--keep] [--transparent]
//
// page   path under public/, e.g. ledger/ (served by tools/serve.py on :8760; started if needed)
// steps  [{ "js": "...", "wait": ms, "shot": "a.jpg" }, ...]
//        js is an async function body evaluated in the page; its return value is printed.
//        shot is written to --out (default /tmp/shots); .png or .jpg by extension.
// --qs     query string appended to the page URL
// --ready  JS expression to wait for (default: the document has loaded; then 500 ms for the app to start)
// --keep   leave the tab open (brought to the front) instead of closing it
// --transparent  no default white page background, so .png shots keep transparency (e.g. icons)
// --inject a.js,b.js  evaluate helper scripts in the page once it is ready
// --timeout ms      longest a single step (or CDP call) may take, default 120000. A hung or crashed
//                   page still gets its tab closed (also on Ctrl-C / SIGTERM).
// Env: CDP (default http://127.0.0.1:9222), BASE (default http://127.0.0.1:8760)
import fs from 'fs';
import path from 'path';
import net from 'net';
import { spawn } from 'child_process';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const argv = process.argv.slice(2);
const opt = (name, def) => { const i = argv.indexOf(name); return i < 0 ? def : argv.splice(i, 2)[1]; };
const flag = (name) => { const i = argv.indexOf(name); return i >= 0 && !!argv.splice(i, 1); };
const W = +opt('--w', 1280), H = +opt('--h', 800), qs = opt('--qs', '');
const ready = opt('--ready', "document.readyState === 'complete'");
const out = opt('--out', '/tmp/shots');
const inject = opt('--inject', '');
const stepMs = +opt('--timeout', 120000);
const keep = flag('--keep'), transparent = flag('--transparent');
const [page = '', stepsArg = '[]'] = argv;
const steps = stepsArg.trim().startsWith('[') ? JSON.parse(stepsArg) : JSON.parse(fs.readFileSync(stepsArg, 'utf8'));
const CDP = process.env.CDP || 'http://127.0.0.1:9222';
const BASE = process.env.BASE || 'http://127.0.0.1:8760';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
fs.mkdirSync(out, { recursive: true });

// start the no-cache server if nothing listens on BASE's port
const port = +new URL(BASE).port;
const listening = () => new Promise((r) => { const s = net.connect(port, '127.0.0.1', () => { s.end(); r(true); }).on('error', () => r(false)); });
let server = null;
if (!(await listening())) {
  server = spawn('python3', [path.join(root, 'tools/serve.py'), String(port)], { stdio: 'ignore' });
  for (let i = 0; i < 50 && !(await listening()); i++) await sleep(100);
}

const target = await (await fetch(`${CDP}/json/new?about:blank`, { method: 'PUT' })).json();
let closed = false;
const closeTab = async () => { if (closed || keep) return; closed = true; await fetch(`${CDP}/json/close/${target.id}`).catch(() => {}); };
for (const sig of ['SIGINT', 'SIGTERM']) process.on(sig, async () => { await closeTab(); server?.kill(); process.exit(130); });
const ws = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((r, j) => { ws.onopen = r; ws.onerror = j; });
let seq = 0;
const pending = new Map(), logs = [];
ws.onmessage = (e) => {
  const m = JSON.parse(e.data);
  if (m.id && pending.has(m.id)) { pending.get(m.id).res(m); pending.delete(m.id); return; }
  const p = m.params;
  if (m.method === 'Network.requestWillBeSent' && !/^(http:\/\/127\.0\.0\.1|data:|blob:|about:)/.test(p.request.url)) logs.push(`[EXTERNAL] ${p.request.url}`);
  if (m.method === 'Runtime.exceptionThrown') logs.push(`[exception] ${p.exceptionDetails.exception?.description ?? p.exceptionDetails.text}`);
  if (m.method === 'Runtime.consoleAPICalled' && /error|warn/.test(p.type)) logs.push(`[${p.type}] ${p.args.map((a) => a.value ?? a.description).join(' ')}`);
  if (m.method === 'Log.entryAdded' && /error|warning/.test(p.entry.level)) logs.push(`[${p.entry.level}] ${p.entry.text} ${p.entry.url ?? ''}`);
};
// every call fails after stepMs instead of waiting forever on a hung page; a closed socket (crashed
// tab) fails whatever is still pending
const send = (method, params = {}, ms = stepMs) => new Promise((res, rej) => {
  const id = ++seq;
  const t = setTimeout(() => { pending.delete(id); rej(new Error(`${method} timed out after ${ms / 1000} s`)); }, ms);
  pending.set(id, { res: (m) => { clearTimeout(t); res(m); }, rej: (e) => { clearTimeout(t); rej(e); } });
  ws.send(JSON.stringify({ id, method, params }));
});
ws.onclose = () => { for (const p of pending.values()) p.rej(new Error('the page closed or crashed')); pending.clear(); };
const evaluate = async (expression) => {
  const r = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true, userGesture: true });
  if (r.result?.exceptionDetails) throw new Error(r.result.exceptionDetails.exception?.description ?? r.result.exceptionDetails.text);
  return r.result?.result?.value;
};

let failed = null;
try {
  for (const d of ['Page', 'Runtime', 'Network', 'Log']) await send(`${d}.enable`);
  await send('Network.setCacheDisabled', { cacheDisabled: true });
  // the page keeps focus while the browser window is in the background
  await send('Emulation.setFocusEmulationEnabled', { enabled: true });
  await send('Emulation.setDeviceMetricsOverride', { width: W, height: H, deviceScaleFactor: 1, mobile: false });
  if (transparent) await send('Emulation.setDefaultBackgroundColorOverride', { color: { r: 0, g: 0, b: 0, a: 0 } });

  const t0 = Date.now();
  await send('Page.navigate', { url: `${BASE}/${page}${qs ? '?' + qs : ''}` });
  let ok = false;
  while (!ok && Date.now() - t0 < 240000) {
    await sleep(250);
    ok = await evaluate(`!!(${ready})`).catch((e) => { if (/crashed|timed out/.test(e.message)) throw e; return false; });
  }
  if (ok) await sleep(500);
  console.log(ok ? `ready in ${((Date.now() - t0) / 1000).toFixed(1)} s` : 'TIMEOUT waiting for ready');
  for (const f of inject ? inject.split(',') : []) await evaluate(fs.readFileSync(f, 'utf8'));

  for (const st of steps) {
    if (st.js) {
      try {
        const v = await evaluate(`(async () => { ${st.js} })()`);
        if (v !== undefined) console.log('js ->', JSON.stringify(v).slice(0, 50000));
      } catch (e) {
        console.log('js error', e.message);
        if (/crashed|timed out/.test(e.message)) throw e;     // the page is gone or stuck: stop here
      }
    }
    if (st.wait) await sleep(st.wait);
    if (st.shot) {
      const png = st.shot.endsWith('.png');
      const s = await send('Page.captureScreenshot', png ? { format: 'png' } : { format: 'jpeg', quality: 85 });
      fs.writeFileSync(path.join(out, st.shot), Buffer.from(s.result.data, 'base64'));
      console.log('shot', path.join(out, st.shot));
    }
  }
} catch (e) {
  failed = e;
  console.log('ABORTED:', e.message);
} finally {
  for (const l of logs.slice(-60)) console.log(l.slice(0, 600));
  if (!logs.length) console.log('no errors, no external requests');
  if (keep) await send('Page.bringToFront', {}, 5000).catch(() => {}); else await closeTab();
  ws.close();
  server?.kill();
}
if (failed) process.exit(1);
