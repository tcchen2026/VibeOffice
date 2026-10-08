// Own one CDP tab. Never enumerate or close the user's tabs.
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
export async function ensureServer(base = process.env.BASE || 'http://127.0.0.1:8760') {
  if (await fetch(base).then(r => r.ok).catch(() => false)) return () => {};
  const server = spawn('python3', [fileURLToPath(new URL('../serve.py', import.meta.url)), new URL(base).port || '8760'], { stdio: 'ignore' });
  for (let i = 0; i < 60; i++) {
    if (await fetch(base).then(r => r.ok).catch(() => false)) return () => server.kill();
    await pause(100);
  }
  server.kill();
  throw new Error('The local test server did not start');
}
export async function openPage(app, options = {}) {
  const cdp = process.env.CDP || 'http://127.0.0.1:9222';
  const base = process.env.BASE || 'http://127.0.0.1:8760';
  const target = await (await fetch(cdp + '/json/new?about:blank', { method: 'PUT' })).json();
  const ws = new WebSocket(target.webSocketDebuggerUrl), pending = new Map(), errors = [], requests = [];
  let seq = 0, closed = false;
  await new Promise((resolve, reject) => { ws.onopen = resolve; ws.onerror = reject; });
  ws.onmessage = event => {
    const m = JSON.parse(event.data), p = m.params;
    if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); }
    if (m.method === 'Runtime.exceptionThrown') errors.push(p.exceptionDetails.exception?.description || p.exceptionDetails.text);
    if (m.method === 'Runtime.consoleAPICalled' && p.type === 'error') errors.push(p.args.map(a => a.value ?? a.description).join(' '));
    if (m.method === 'Network.requestWillBeSent' && !/^(http:\/\/127\.0\.0\.1(?=[:/])|data:|blob:|about:)/.test(p.request.url)) requests.push(p.request.url);
  };
  const send = (method, params = {}) => new Promise((resolve, reject) => {
    const id = ++seq;
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(method + ' timed out')); }, options.timeout || 120000);
    pending.set(id, msg => { clearTimeout(timer); if (msg.error) reject(new Error(msg.error.message)); else resolve(msg.result); });
    ws.send(JSON.stringify({ id, method, params }));
  });
  const evaluate = async expression => {
    const r = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true, userGesture: true });
    if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || r.exceptionDetails.text);
    return r.result?.value;
  };
  const close = async () => {
    if (closed) return;
    closed = true;
    for (const done of pending.values()) done({ error: { message: 'Test tab closed' } });
    pending.clear();
    await fetch(cdp + '/json/close/' + target.id).catch(() => {});
    ws.close();
  };
  try {
    for (const domain of ['Page', 'Runtime', 'Network']) await send(domain + '.enable');
    await send('Network.setCacheDisabled', { cacheDisabled: true });
    // Capture the actual app snapshot/recovery hooks, without touching the user's IndexedDB.
    await send('Page.addScriptToEvaluateOnNewDocument', { source: `
      let vo;
      Object.defineProperty(window, 'VO', { configurable: true, get: () => vo, set(value) {
        vo = value;
        let attach;
        Object.defineProperty(value, 'attach', { configurable: true, get: () => attach, set(fn) {
          attach = function(app, hooks) { window.__corpusHooks = hooks; return ${options.persistence === false ? 'undefined' : 'fn(app, hooks)'}; };
        }});
      }});` });
    await send('Page.navigate', { url: base + '/' + app + '/?new' });
    for (let i = 0; i < 200; i++) {
      if (await evaluate('!!window.__corpusHooks').catch(() => false)) return { evaluate, close, errors, requests, send };
      await pause(100);
    }
    throw new Error(app + ' did not attach to the suite');
  } catch (e) { await close(); throw e; }
}
