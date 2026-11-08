// VO.host (public/common/suite.js), the pages' side of the desktop app's host interface (docs/desktop.md):
// with the desktop host (tests/desktop/host-stub.js) a file Windows starts us with opens in its app and
// no service worker is registered, the version is the app's and links go to the system; in a browser the
// host is 'web'. The update policies (updater/current.json) and what is new (updater/whatsnew.json); versions
// compare by date (VO.compareVersions).
//
//   node --test tests/desktop/host.test.mjs   (needs Chromium with CDP at 127.0.0.1:9222, as tools/shot.mjs)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '../..');
const shot = (page, steps, ...args) => {
  const out = execFileSync('node', [path.join(root, 'tools/shot.mjs'), page, JSON.stringify(steps), ...args], { encoding: 'utf8' });
  const line = out.split('\n').findLast((l) => l.startsWith('js -> '));   // the last step's result
  assert.ok(line, out);
  return { out, r: JSON.parse(JSON.parse(line.slice(6))) };
};
// the stub's records (window.stubSession); in a plain browser, its sessionStorage
const S = '(window.stubSession || sessionStorage)';
const state = `return JSON.stringify({ path: location.pathname, title: document.title, kind: VO.host.kind,
  platform: VO.host.platform, version: VO.VERSION, newWindow: ${S}.newWindow || '', calls: ${S}.hostCalls || '', touched: ${S}.touched || '' });`;

const stubbed = ['--pre', path.join(root, 'tests/desktop/host-stub.js')];
const wait = (ms) => `await new Promise((r) => setTimeout(r, ${ms}));`;

test('desktop: a window made for a file opens it, from the native side', () => {
  const { out, r } = shot('ledger/', [{ wait: 2500 }, { js: state }], ...stubbed, '--qs', 'doc&file=Budget.csv');
  assert.equal(r.path, '/ledger/');
  assert.match(r.title, /^Budget\.csv/);
  assert.equal(r.kind, 'desktop');
  assert.equal(r.platform, 'windows');
  assert.match(r.calls, /\/ledger\/:host_document .*\/ledger\/:host_document_bytes/, 'asked what it holds, then took it');
  assert.equal(r.touched, '', 'no browser storage, no service worker in the desktop app');
  assert.equal(r.version, 'v2026.10.11-abc123', 'the desktop build\'s version');
  assert.match(out, /no errors, no external requests/);
});

test('desktop: Open File in the Start Center asks the system\'s Open dialog, and the file opens in its app', () => {
  const { out, r } = shot('', [{ js: `document.getElementById('open').click(); return '1'`, wait: 2500 },
    { js: `return JSON.stringify({ path: location.pathname, title: document.title, newWindow: stubSession.newWindow, dialogs: JSON.parse(stubSession.dialogs) })` }],
  ...stubbed, '--qs', 'pick=Notes.md');
  assert.equal(r.path, '/quire/', 'Notes.md went to Quire');
  assert.match(r.newWindow, /\/quire\/\?doc$/, 'in a window of its own, made for it');
  assert.match(r.title, /^Notes/);
  assert.deepEqual(r.dialogs, { title: 'Open', all: 'All files VibeOffice opens', quire: 'Documents', ledger: 'Spreadsheets', lectern: 'Presentations' });
  assert.match(out, /no errors, no external requests/);
});

test('desktop: File > Open asks to save a changed document, then the window reloads with the file; the page it replaces gets nothing', () => {
  const ask = (answer) => `L.D.doc.dirty = true; const p = L.app.openDialog(); ${wait(500)} const asked = /save the changes/.test(${boxText});
    ${click(answer)} await p;
    /* the page being replaced asks for the window's new document before the reload: it doesn't know the word */
    let info = 'error', bytes = 'given';
    try { info = await window.__TAURI_INTERNALS__.invoke('host_document', {}); } catch (e) { /* refused */ }
    try { await window.__TAURI_INTERNALS__.invoke('host_document_bytes', {}); } catch (e) { bytes = 'refused'; }
    stubSession.old = JSON.stringify({ asked, info, bytes, calls: stubSession.hostCalls.split(' ').map((c) => c.split(':').pop()).filter((c) => /reopen|clear_draft/.test(c)) });
    return '1'`;
  const then = `return JSON.stringify({ old: JSON.parse(stubSession.old), path: location.pathname, title: document.title, reopened: stubSession.reopened || '', newWindow: stubSession.newWindow || '' })`;
  const no = shot('quire/', [{ wait: 1500 }, { js: ask('No'), wait: 2500 }, { js: then }], ...stubbed, '--qs', 'new&pick=Notes.md').r;
  assert.deepEqual(no.old, { asked: true, info: null, bytes: 'refused', calls: ['host_clear_draft', 'host_reopen'] }, 'No gives the changes up; the old page is given nothing');
  assert.match(no.reopened, /^\/quire\/\?doc&t=\w+$/);
  assert.deepEqual([no.path, no.title, no.newWindow], ['/quire/', 'Notes - Quire', ''], 'the same window, reloaded with the file');
  const cancel = shot('quire/', [{ wait: 1500 }, { js: ask('Cancel'), wait: 1500 }, { js: then }], ...stubbed, '--qs', 'new&pick=Notes.md').r;
  assert.deepEqual([cancel.old.asked, cancel.old.calls, cancel.reopened, cancel.title], [true, ['host_reopen'], '', 'Document1 - Quire'], 'Cancel keeps the document');
});

test('desktop: an app window\'s Recent Files are the Start Center\'s: its lists lead there, and it never reads them', () => {
  const { r } = shot('quire/', [{ wait: 2500 }, { js: `const items = VO.recentMenuItems(4); items[0].run(); ${wait(300)}
    const links = [...VO.recentLinks(4).querySelectorAll('button')].map((b) => b.textContent);
    return JSON.stringify({ items: items.map((i) => i.label), links, calls: stubSession.hostCalls.split(' ').map((c) => c.split(':').pop()), touched: stubSession.touched || '' })` }],
  ...stubbed, '--qs', 'doc&file=Notes.md&recent=Budget.csv,Notes.md*');
  assert.deepEqual(r.items, ['Recent &Files...']);
  assert.deepEqual(r.links, ['Recent Files...']);
  assert.ok(r.calls.includes('host_show_recent'), r.calls.join(' '));
  assert.ok(!r.calls.some((c) => /host_recent|host_open_recent|host_discard_draft/.test(c)), 'no list, no entry opened by key: ' + r.calls.join(' '));
});

test('desktop: a file dropped on a window with a document opens in a new window, with its bytes', () => {
  const { r } = shot('ledger/', [{ wait: 2000 }, { js: `L.app.openFiles([new File(['x,y\\n1,2\\n'], 'Dropped.csv')]); return '1'`, wait: 2500 },
    { js: `return JSON.stringify({ path: location.pathname, title: document.title, newWindow: stubSession.newWindow || '' })` }], ...stubbed, '--qs', 'doc&file=Budget.csv');
  assert.deepEqual(r, { path: '/ledger/', title: 'Dropped.csv - Ledger', newWindow: 'http://127.0.0.1:8760/ledger/?doc' });
});

test('desktop: the Start Center lists Recent Files the native side keeps, and opens an unsaved version', () => {
  const names = `[...document.querySelectorAll('#grid .doc')].map((c) => c.querySelector('.name').textContent + (c.querySelector('.unsaved') ? ' (unsaved)' : '') + (c.querySelector('img') ? ' [preview]' : ''))`;
  const { r } = shot('', [{ wait: 1500 }, { js: `stubSession.cards = JSON.stringify(${names}); document.querySelector('#grid .rm').click(); ${wait(300)}
    stubSession.after = JSON.stringify(${names}); document.querySelector('#grid .doc .name').click(); return '1'`, wait: 2500 },
  { js: `return JSON.stringify({ cards: JSON.parse(stubSession.cards), after: JSON.parse(stubSession.after), path: location.pathname, recovered: document.body.textContent.includes('Recovered the unsaved changes') })` }],
  ...stubbed, '--qs', 'recent=Other.csv,Budget.csv*');
  assert.deepEqual(r, { cards: ['Other.csv [preview]', 'Budget.csv (unsaved) [preview]'], after: ['Budget.csv (unsaved) [preview]'], path: '/ledger/', recovered: true },
    '× removed Other.csv; Budget.csv opened at its unsaved version');
});

test('desktop: the close button asks the page, which asks to save its changes', () => {
  const ask = (answer) => `L.D.doc.dirty = true; window.dispatchEvent(new Event('vo-close')); ${wait(400)}
    const box = document.querySelector('.msgbox, dialog[open], .dlg'); const text = box ? box.textContent : '';
    [...document.querySelectorAll('button')].find((b) => b.textContent.trim() === '${answer}').click(); ${wait(400)}
    return JSON.stringify({ asked: /save the changes/.test(text), closed: stubSession.closed || '' })`;
  assert.deepEqual(shot('quire/', [{ wait: 1500 }, { js: ask('Cancel') }], ...stubbed, '--qs', 'new').r, { asked: true, closed: '' }, 'Cancel keeps it');
  assert.deepEqual(shot('quire/', [{ wait: 1500 }, { js: ask('No') }], ...stubbed, '--qs', 'new').r, { asked: true, closed: '/quire/' }, 'No closes it');
});

test('desktop: Esc in "Save changes?" cancels closing, in every app: the document, its changes and its draft stay', () => {
  const dirty = { quire: 'L.D.doc.dirty = true', ledger: 'L.app.books[L.app.cur].dirty = true', lectern: 'L.hist.dirty = true' };
  const still = { quire: 'L.D.doc.dirty', ledger: 'L.app.books[L.app.cur].dirty', lectern: 'L.hist.dirty' };
  for (const app of ['quire', 'ledger', 'lectern']) {
    const { r } = shot(app + '/', [{ wait: 2000 }, { js: `${dirty[app]}; stubSession.hostCalls = ''; window.dispatchEvent(new Event('vo-close')); ${wait(400)}
      const box = document.querySelector('.msgbox'); const text = box ? box.textContent : '';
      (document.activeElement || document.body).dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })); ${wait(400)}
      return JSON.stringify({ asked: /save the changes/.test(text), open: !!document.querySelector('.msgbox'), closed: stubSession.closed || '',
        calls: stubSession.hostCalls.split(' ').map((c) => c.split(':').pop()).filter((c) => /close_window|clear_draft|discard/.test(c)), dirty: !!(${still[app]}) })` }],
    ...stubbed, '--qs', 'new');
    assert.deepEqual(r, { asked: true, open: false, closed: '', calls: [], dirty: true }, app);
  }
});

test('autosave: a draft 15 s after the last change, then at most every 2 minutes, a large document less often; typing on doesn\'t put it off; over 50 MB none, said once', () => {
  const { out, r } = shot('quire/', [{ wait: 2000 }, { js: `const doc = L.D.doc, realTimeout = window.setTimeout, realNow = Date.now;
    let now = 0, fire = null;
    Date.now = () => now;
    /* at t seconds, a change: how long until its draft (s); the timer is kept to be fired by hand */
    const change = (t) => { now = t * 1000; let ms = null; window.setTimeout = (fn, d) => { fire = fn; ms = d; return 0; };
      try { doc.dirty = true; VO.changed(doc); } finally { window.setTimeout = realTimeout; } return ms / 1000; };
    /* the draft written at t seconds (bytes: a document of that size), its size as the host got it */
    const write = async (t, bytes) => { now = t * 1000; stubSession.draft = ''; if (bytes) L.docx.write = async () => new Blob([new Uint8Array(bytes)]);
      await fire(); ${wait(300)} return stubSession.draft ? JSON.parse(stubSession.draft).size : 0; };
    const toast = () => { const e = document.querySelector('.toast'); const shown = !!e && !e.hidden && /over 50 MB/.test(e.textContent); if (e) e.hidden = true; return shown; };
    const waits = [change(0), change(5), change(110)];
    const first = await write(120) > 0;
    waits.push(change(130));
    const mb24 = await write(240, 24 << 20);
    waits.push(change(250));
    const mb51 = await write(528, 51 << 20), told = toast();
    waits.push(change(530));
    await write(1140, 51 << 20);
    const toldAgain = toast();
    Date.now = realNow;
    return JSON.stringify({ waits, first, mb24, mb51, told, toldAgain });` }], ...stubbed, '--qs', 'new');
  assert.deepEqual(r.waits, [15, 15, 10, 110, 278, 610],
    '15 s after the last change; typing on, 2 minutes after the first; then 2 minutes after the last draft; 12 s a MB after 24 MB; after 51 MB too');
  assert.deepEqual([r.first, r.mb24, r.mb51], [true, 24 << 20, 0], 'the drafts the host got: none over 50 MB');
  assert.deepEqual([r.told, r.toldAgain], [true, false], 'a document over 50 MB says so once');
  assert.match(out, /no errors, no external requests/);
});

test('desktop: the Mac\'s menu bar runs Open… and Help in the window in front (vo-menu), never behind its dialog', () => {
  const menu = (what) => `window.dispatchEvent(new CustomEvent('vo-menu', { detail: '${what}' })); ${wait(400)}`;
  const asked = `!!stubSession.dialogs`;
  const { r } = shot('quire/', [{ wait: 1800 }, { js: `${menu('help')} const help = document.querySelector('#taskpane .tp-title').textContent;
    L.ui.exec('about'); ${wait(300)} ${menu('open')} const behind = ${asked};
    document.querySelector('.dlg .btn.primary').click(); ${wait(300)} ${menu('open')}
    return JSON.stringify({ help, behind, open: ${asked} })` }], ...stubbed, '--qs', 'new');
  assert.deepEqual(r, { help: 'Quire Help', behind: false, open: true }, 'Help shows the app\'s Help; Open… waits for About to close');
  assert.equal(shot('', [{ wait: 1500 }, { js: `${menu('open')} return JSON.stringify(${asked})` }], ...stubbed).r, true, 'the Start Center\'s Open File');
});

test('the stub grants each kind of window the commands the desktop app does (desktop/permissions/, capabilities/)', () => {
  const list = (text) => [...text.matchAll(/["'](host_\w+)["']/g)].map((m) => m[1]).sort();
  const perms = {};
  for (const m of fs.readFileSync(path.join(root, 'desktop/permissions/host.toml'), 'utf8').matchAll(/identifier = "([^"]+)"[\s\S]*?commands\.allow = \[([^\]]*)\]/g)) perms[m[1]] = list(m[2]);
  const caps = {};
  for (const f of fs.readdirSync(path.join(root, 'desktop/capabilities'))) {
    const c = JSON.parse(fs.readFileSync(path.join(root, 'desktop/capabilities', f), 'utf8'));
    caps[c.identifier + ': ' + c.windows.join(' ')] = c.permissions.flatMap((p) => perms[p]).sort();
  }
  const stub = fs.readFileSync(path.join(root, 'tests/desktop/host-stub.js'), 'utf8');
  const grants = (kind) => list(new RegExp(`\\b${kind}: \\[([^\\]]*)\\]`).exec(stub)[1]);
  /* lite: the lite build's one window instead of start and apps (its TAURI_CONFIG; desktop/tauri.conf.json) */
  assert.deepEqual(caps, { 'start: main': grants('start'), 'apps: quire-* ledger-* lectern-*': grants('app'), 'lite: main': grants('lite') });
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(root, 'desktop/tauri.conf.json'), 'utf8')).app.security.capabilities, ['start', 'apps']);
});

test('no browser storage in the desktop app: the stub removes what host.rs does, and only the web\'s own code names it', () => {
  const pairs = (text) => [...text.matchAll(/\[(window|navigator), '(\w+)'\]/g)].map((m) => m[1] + '.' + m[2]);
  const rust = /NO_STORAGE: &str = ([\s\S]*?";)/.exec(fs.readFileSync(path.join(root, 'desktop/src/host.rs'), 'utf8'))[1];
  assert.deepEqual(pairs(fs.readFileSync(path.join(root, 'tests/desktop/host-stub.js'), 'utf8')), pairs(rust));
  assert.equal(pairs(rust).length, 7);
  /* the web branches of the settings (common/look.js), Recent Files and the service worker (suite.js) and the
     Start Center (start.js); sw.js is the service worker itself */
  const allowed = new Set(['common/look.js', 'common/suite.js', 'start.js', 'sw.js']);
  const names = /\b(localStorage|sessionStorage|indexedDB|BroadcastChannel|navigator\.storage|navigator\.serviceWorker|caches\.(open|match|keys|delete|has))\b/;
  const bad = [];
  const walk = (dir) => {
    for (const e of fs.readdirSync(path.join(root, 'public', dir), { withFileTypes: true })) {
      const rel = dir ? dir + '/' + e.name : e.name;
      if (e.isDirectory()) { if (rel !== 'common/i18n') walk(rel); }
      else if (/\.(js|html)$/.test(e.name) && !allowed.has(rel) && names.test(fs.readFileSync(path.join(root, 'public', rel), 'utf8'))) bad.push(rel);
    }
  };
  walk('');
  assert.deepEqual(bad, []);
});

test('desktop: the settings are the host\'s: read as a page starts (after a reload too), kept by the host, changed by another window; no page uses browser storage', () => {
  const { r } = shot('', [
    { js: `VO.look.set('classic'); document.getElementById('settings').click(); await new Promise((r) => setTimeout(r, 300));
      const s = document.getElementById('uilang'); s.value = 'fr'; s.dispatchEvent(new Event('change')); return '1'`, wait: 2500 },
    { js: `stubSession.reloaded = JSON.stringify({ lang: VO.lang.get(), look: document.documentElement.dataset.look, how: performance.getEntriesByType('navigation')[0].type });
      setTimeout(() => { location.href = '/quire/?new'; }); return '1'`, wait: 3500 },
    { js: `const started = { lang: document.documentElement.lang, look: document.documentElement.dataset.look };
      L.store.set('testKey', { a: 1 });
      window.dispatchEvent(new CustomEvent('vo-settings', { detail: { key: 'vibeoffice.look', value: 'paper2016' } }));
      await new Promise((r) => setTimeout(r, 200));
      const kept = JSON.parse(stubSession.stubSettings);
      return JSON.stringify({ reloaded: JSON.parse(stubSession.reloaded), started, kept: [kept['vibeoffice.look'], kept['vibeoffice.lang'], kept['quire.testKey']],
        read: L.store.get('testKey'), changed: [VO.settings.get('vibeoffice.look'), document.documentElement.dataset.layout], touched: stubSession.touched || '' })` }], ...stubbed);
  // (tools/shot.mjs shows the Start Center in English, ?lang=en; the language picked is what it read)
  assert.deepEqual(r.reloaded, { lang: 'fr', look: 'classic', how: 'reload' }, 'the Start Center reloads with the language just picked');
  assert.deepEqual(r.started, { lang: 'fr', look: 'classic' }, 'a new window starts with the settings');
  assert.deepEqual(r.kept, ['classic', 'fr', '{"a":1}'], 'the host keeps the look, the language and L.store\'s values');
  assert.deepEqual(r.read, { a: 1 });
  assert.deepEqual(r.changed, ['paper2016', 'ribbon'], 'another window\'s look applies here');
  assert.equal(r.touched, '', 'no page used browser storage');
});

test('desktop: File > Exit closes an app window', () => {
  const { r } = shot('quire/', [{ js: `await VO.exit(); await new Promise((r) => setTimeout(r, 300));
    return JSON.stringify({ closed: stubSession.closed || '', path: location.pathname });` }], ...stubbed);
  assert.deepEqual(r, { closed: '/quire/', path: '/quire/' });
});

test('desktop: About > Open source licenses opens the credits page in a window of its own', () => {
  const { r } = shot('', [{ js: `document.getElementById('about').click(); document.getElementById('osslink').click(); return '1'`, wait: 800 },
    { js: `return JSON.stringify({ newWindow: stubSession.newWindow || '', title: document.title })` }],
  ...stubbed);
  assert.match(r.newWindow, /\/oss-licenses\.html$/);
  assert.match(r.title, /^Open source licenses/);
});

test('desktop: a link on the licences page shows its address to copy, and goes nowhere', () => {
  const { r } = shot('oss-licenses.html', [{ js: `document.querySelector('.head a').click(); await new Promise((r) => setTimeout(r, 200));
    const d = document.querySelector('dialog');
    return JSON.stringify({ open: d.open, address: d.querySelector('input').value, href: document.querySelector('.head a').href, path: location.pathname });` }],
  ...stubbed);
  assert.deepEqual(r, { open: true, address: r.href, href: r.href, path: '/oss-licenses.html' });
  assert.match(r.href, /^https:/);
});

test('desktop: a window that can\'t be opened is reported, and the Start Center stays', () => {
  const { r } = shot('', [{ js: `let error = ''; try { await VO.go(location.origin + '/updater/'); } catch (e) { error = String(e); }
    await new Promise((r) => setTimeout(r, 300)); return JSON.stringify({ error, path: location.pathname, newWindow: stubSession.newWindow || '' });` }],
  ...stubbed);
  assert.deepEqual(r, { error: "not one of the suite's pages: updater", path: '/', newWindow: '' });
});

const saves = `JSON.parse(stubSession.saves || '[]')`;
const click = (label) => `[...document.querySelectorAll('.dlg button, .msgbox button')].find((b) => b.textContent.replace(/\\s+/g, ' ').trim() === '${label}').click();`;
const boxText = `[...document.querySelectorAll('.dlg')].map((d) => d.textContent).join(' ')`;

test('desktop: Save writes the window\'s own file in place, without asking', () => {
  const { r } = shot('quire/', [{ wait: 2500 }, { js: `L.D.doc.dirty = true; const r = await L.app.save();
    return JSON.stringify({ r, saves: ${saves}, chose: stubSession.chose || '', dirty: L.D.doc.dirty, title: document.title })` }], ...stubbed, '--qs', 'doc&file=Notes.md');
  assert.deepEqual(r.saves.map((w) => [w.cmd, w.replace, w.head]), [['host_save', false, '# ']], 'one write, Markdown into Notes.md');
  assert.deepEqual({ r: r.r, chose: r.chose, dirty: r.dirty, title: r.title }, { r: 'saved', chose: '', dirty: false, title: 'Notes - Quire' });
});

test('desktop: a document without a file asks where, and the file it is saved as becomes its own', () => {
  const { r } = shot('quire/', [{ js: `stubSession.stubSave = 'Letter.docx'; L.D.doc.dirty = true; const first = await L.app.save();
    const title = document.title; L.D.doc.dirty = true; const second = await L.app.save();
    return JSON.stringify({ first, second, title, saves: ${saves}, chose: JSON.parse(stubSession.chose) })` }], ...stubbed, '--qs', 'new');
  assert.deepEqual([r.first, r.second, r.title], ['saved', 'saved', 'Letter - Quire']);
  assert.deepEqual(r.saves.map((w) => w.cmd + (w.bound ? ' (bound)' : '')), ['host_write_chosen (bound)', 'host_save'], 'Save As, then Save in place');
  assert.equal(r.chose.purpose, 'saveAs');
  assert.equal(r.chose.type, 'docx');
  assert.equal(r.chose.name, 'Document1');
  assert.equal(r.chose.types.length, 10, 'every type Quire saves');
});

test('desktop: changes made while a save runs stay unsaved', () => {
  const { r } = shot('quire/', [{ wait: 2500 }, { js: `const d = L.D.doc; d.dirty = true; const p = L.app.save();
    VO.changed(d); d.dirty = true; const r = await p; return JSON.stringify({ r, dirty: d.dirty })` }], ...stubbed, '--qs', 'doc&file=Notes.md');
  assert.deepEqual(r, { r: 'saved', dirty: true });
});

test('desktop: Ledger saves the CSV it opened, and is then saved; its recovered draft (a workbook) goes back as CSV', () => {
  const { r } = shot('ledger/', [{ wait: 2500 }, { js: `const b = L.app.books[L.app.cur]; b.dirty = true;
      stubSession.first = JSON.stringify({ saved: await L.app.save(), dirty: b.dirty, saves: ${saves}.map((s) => s.cmd + (s.head === 'PK' ? ' zip' : ' text')) });
      stubSession.saves = '[]';
      const u8 = new Uint8Array(await (await L.xlsxWrite.write(b.wb, { type: 'xlsx' })).arrayBuffer());
      let s = ''; for (const c of u8) s += String.fromCharCode(c);
      stubSession.stubDoc = JSON.stringify({ key: 'k-Budget.csv', b64: btoa(s),
        info: { fileName: 'Budget.xlsx', name: 'Budget.csv', saved: true, draft: true, lossState: null, bound: true, readOnly: false } });
      setTimeout(() => { location.href = '/ledger/?doc'; }); return '1'`, wait: 3000 },
    { js: `const b = L.app.books[L.app.cur]; const before = { type: b.type, dirty: b.dirty, title: document.title };
      return JSON.stringify({ first: JSON.parse(stubSession.first), before, saved: await L.app.save(), dirty: b.dirty, saves: ${saves}.map((s) => s.cmd + (s.head === 'PK' ? ' zip' : ' text')) })` }],
  ...stubbed, '--qs', 'doc&file=Budget.csv');
  assert.deepEqual(r.first, { saved: true, dirty: false, saves: ['host_save text'] }, 'Save writes the CSV in place, and closing has nothing to ask');
  assert.deepEqual([r.before.type, r.before.dirty], ['csv', true], 'the recovered draft is the CSV file\'s');
  assert.deepEqual({ saved: r.saved, dirty: r.dirty, saves: r.saves }, { saved: true, dirty: false, saves: ['host_save text'] }, 'saved as CSV text, not the draft\'s workbook');
});

test('desktop: a read-only file says so in the title, and Save asks where', () => {
  const { r } = shot('quire/', [{ wait: 2500 }, { js: `const title = document.title; L.D.doc.dirty = true; const r = await L.app.save();
    return JSON.stringify({ title, r, chose: JSON.parse(stubSession.chose || 'null') })` }], ...stubbed, '--qs', 'doc&file=Notes.md&readonly');
  assert.equal(r.title, 'Notes (Read-Only) - Quire');
  assert.equal(r.r, 'declined', 'the Save dialog was cancelled');
  assert.equal(r.chose && r.chose.purpose, 'saveAs');
});

test('desktop: a save the Mac sandbox refuses asks for the file\'s folder, then saves; elsewhere a refusal offers Save As', () => {
  const run = (qs, answer) => shot('quire/', [{ wait: 2500 }, { js: `L.D.doc.dirty = true; const p = L.app.save(); ${wait(500)} const text = ${boxText};
    const buttons = [...document.querySelectorAll('.msgbox button, .dlg button')].map((b) => b.textContent.replace(/\\s+/g, ' ').trim()).filter(Boolean);
    ${click(answer)} const r = await p; return JSON.stringify({ r, text, buttons, allowed: JSON.parse(stubSession.allowed || 'null'), saves: ${saves}.map((w) => w.cmd) })` }],
  ...stubbed, '--qs', 'doc&file=Notes.md&fail=denied' + qs).r;
  const mac = run('&sandboxed', 'Allow Access...');
  assert.match(mac.text, /doesn't have permission to save “Notes\.md” in its folder/);
  assert.deepEqual(mac.buttons.slice(-3), ['Allow Access...', 'Save As...', 'Cancel']);
  assert.equal(mac.allowed.how, 'save', 'the folder of the window\'s own file');
  assert.match(mac.allowed.text, /Choose the folder of “Notes\.md”/);
  assert.deepEqual([mac.r, mac.saves], ['saved', ['host_save']], 'saved once allowed');
  const other = run('', 'Cancel');
  assert.match(other.text, /it is read-only, or you don't have permission to change it/);
  assert.deepEqual(other.buttons.slice(-2), ['Save As...', 'Cancel']);
  assert.deepEqual([other.r, other.allowed], ['cancelled', null]);
});

test('desktop: a file another program changed asks first, and Replace It saves over it', () => {
  const { r } = shot('quire/', [{ wait: 2500 }, { js: `L.D.doc.dirty = true; const p = L.app.save(); ${wait(500)} const text = ${boxText};
    ${click('Replace It')} const r = await p; return JSON.stringify({ r, text, saves: ${saves} })` }], ...stubbed, '--qs', 'doc&file=Notes.md&fail=changed');
  assert.match(r.text, /“Notes\.md” was changed by another program after you opened it/);
  assert.deepEqual(r.saves.map((w) => [w.cmd, w.replace]), [['host_save', true]]);
  assert.equal(r.r, 'saved');
});

test('desktop: a PDF export asks where, and leaves the document as it was', () => {
  const { r } = shot('quire/', [{ wait: 2500 }, { js: `stubSession.stubSave = 'Notes.pdf'; L.D.doc.dirty = true; await L.app.printPDF();
    return JSON.stringify({ saves: ${saves}, chose: JSON.parse(stubSession.chose), dirty: L.D.doc.dirty, title: document.title })` }], ...stubbed, '--qs', 'doc&file=Notes.md');
  assert.deepEqual(r.saves.map((w) => [w.cmd, w.type, w.bound]), [['host_write_chosen', 'pdf', false]]);
  assert.deepEqual([r.chose.purpose, r.chose.types.map((t) => t.id)], ['export', ['pdf']]);
  assert.deepEqual([r.dirty, r.title], [true, 'Notes - Quire'], 'still the Markdown file, with its changes');
});

test('Ledger: Cancel in Save As while closing an untitled workbook keeps it open', () => {
  const { r } = shot('ledger/', [{ wait: 1500 }, { js: `const b = L.app.books[0]; b.dirty = true; const p = L.app.closeWorkbook(); ${wait(400)}
    ${click('Yes')} ${wait(400)} ${click('Cancel')} const r = await p;
    return JSON.stringify({ r, same: L.app.books[0] === b, dirty: b.dirty })` }], '--qs', 'new');
  assert.deepEqual(r, { r: false, same: true, dirty: true });
});

test('desktop: openUrl sends the update page and our GitHub to the system, and nothing else', () => {
  const { r } = shot('', [{ js: `await VO.host.openUrl('https://github.com/tcchen2026/VibeOffice');
    const bad = []; for (const url of ['file:///C:/Windows/notepad.exe', 'https://example.com/', 'mailto:someone@example.com', 'https://vibeoffice.work/'])
      try { await VO.host.openUrl(url); } catch (e) { bad.push(String(e)); }
    return JSON.stringify({ opened: stubSession.opened, bad });` }], ...stubbed);
  assert.equal(r.opened, 'https://github.com/tcchen2026/VibeOffice');
  assert.equal(r.bad.length, 4);
  for (const e of r.bad) assert.match(e, /not opened/);
});

const settings = `document.getElementById('settings').click(); await new Promise((r) => setTimeout(r, 400));
  const out = { row: !document.getElementById('scalerow').hidden, hint: !document.getElementById('scaleweb').hidden, value: document.getElementById('uiscale').value };
  if (out.row) { const s = document.getElementById('uiscale'); s.value = '2'; s.dispatchEvent(new Event('change')); await new Promise((r) => setTimeout(r, 200)); out.saved = stubSession.uiScale;
    let bad = ''; try { await VO.host.setUiScale(9); } catch (e) { bad = String(e); } out.bad = bad; }
  return JSON.stringify(out);`;

test('desktop: Settings sets the UI size', () => {
  const { r } = shot('', [{ js: settings }], ...stubbed);
  assert.deepEqual(r, { row: true, hint: false, value: '1', saved: '2', bad: 'RangeError: the UI size is from 0.5 to 4' });
});

test('web: Settings points to the browser\'s zoom', () => {
  const { r } = shot('', [{ js: settings }]);
  assert.deepEqual(r, { row: false, hint: true, value: '1' });
});

const update = `await new Promise((r) => setTimeout(r, 500));
  const out = { button: !document.getElementById('update').hidden, check: !document.getElementById('updatecheck').hidden,
    dialog: document.getElementById('updatedlg').open };
  if (out.dialog) out.news = [...document.querySelectorAll('#updatenews > *')].map((e) => e.textContent);
  if (out.check) out.age = document.getElementById('updateage').textContent;
  if (out.dialog) { const d = document.getElementById('updatedlg'); d.querySelector('button[value=download]').click(); }
  else if (out.button) document.getElementById('update').click();
  else if (out.check) document.getElementById('updatecheck').click();
  await new Promise((r) => setTimeout(r, 200));
  out.opened = ${S}.opened || '';
  VO.settings.remove('vibeoffice.updateAsked');
  return JSON.stringify(out);`;
const stub = path.join(root, 'tests/desktop/host-stub.js');
const URL_ = 'https://vibeoffice.work/updater/?from_version=v2026.10.11-abc123&current_platform=windows';

test('desktop: up to date, no Update button', () => {
  const { r } = shot('', [{ js: update }], '--pre', stub);
  assert.deepEqual(r, { button: false, check: false, dialog: false, opened: '' });
});

test('desktop: show-update-button shows the button, which opens the updater page', () => {
  const { r } = shot('', [{ js: update }], '--pre', stub, '--qs', 'update=button');
  assert.deepEqual(r, { button: true, check: false, dialog: false, opened: URL_ });
});

test('desktop: show-update-dialog asks with a dialog, with what is new', () => {
  const { r } = shot('', [{ js: update }], '--pre', stub, '--qs', 'update=dialog');
  assert.deepEqual(r, { button: true, check: false, dialog: true, opened: URL_,
    news: ['v2026.10.14-000000', 'Fourteenth.', 'v2026.10.12-000000', 'Twelfth.\nSecond line.'] }, 'what is new since this version, newest first');
});

test('the published updater/current.json: a build from 10-08 gets the button, 10-01 the dialog, today\'s nothing', () => {
  for (const [version, want] of [['v2026.10.08-000000', { button: true, dialog: false }], ['v2026.10.01-000000', { button: true, dialog: true }],
    ['v2026.10.11-32d7b2', { button: false, dialog: false }]]) {
    const { r } = shot('', [{ js: update }], '--pre', stub, '--qs', `version=${version}`);
    assert.deepEqual({ button: r.button, dialog: r.dialog }, want, version);
  }
});

test('desktop: a version over two months old offers to check for updates, when current.json asks nothing or cannot be read', () => {
  for (const policy of ['none', 'offline']) {
    const { r } = shot('', [{ js: update }], '--pre', stub, '--qs', `update=${policy}&version=v2026.08.16-000000`);
    assert.deepEqual(r, { button: false, check: true, dialog: false, age: 'This version is 2 months old.',
      opened: 'https://vibeoffice.work/updater/?from_version=v2026.08.16-000000&current_platform=windows' }, policy);
  }
  const { r } = shot('', [{ js: update }], '--pre', stub, '--qs', 'update=none&version=v2026.08.17-000000');
  assert.deepEqual(r, { button: false, check: false, dialog: false, opened: '' }, 'a day short of two months');
});

test('desktop: when current.json asks for an update, an old version gets the Update button, not Check for updates', () => {
  const { r } = shot('', [{ js: update }], '--pre', stub, '--qs', 'update=button&version=v2026.07.01-000000');
  assert.deepEqual({ button: r.button, check: r.check, opened: r.opened }, { button: true, check: false,
    opened: 'https://vibeoffice.work/updater/?from_version=v2026.07.01-000000&current_platform=windows' });
});

test('desktop: the Mac App Store edition never reads current.json; old, its Check for updates opens the page for the App Store', () => {
  const { r } = shot('', [{ js: update }], '--pre', stub, '--qs', 'appstore&update=dialog&version=v2026.07.01-000000');
  assert.deepEqual(r, { button: false, check: true, dialog: false, age: 'This version is 3 months old.',
    opened: 'https://vibeoffice.work/updater/?from_version=v2026.07.01-000000&current_platform=macos_store' }, 'no Update button or dialog, though current.json asks');
  assert.deepEqual(shot('', [{ js: update }], '--pre', stub, '--qs', 'appstore&update=dialog&version=v2026.10.01-000000').r,
    { button: false, check: false, dialog: false, opened: '' }, 'two weeks old');
});

test('web: no update check', () => {
  const { r } = shot('', [{ js: update }]);
  assert.deepEqual(r, { button: false, check: false, dialog: false, opened: '' });
});

test('web and the installed PWA: a copy over two months old offers to check for updates', () => {
  const { r } = shot('', [{ js: update }], '--pre', stub, '--qs', 'web&version=v2026.08.16-000000');
  assert.deepEqual(r, { button: false, check: true, dialog: false, age: 'This version is 2 months old.',
    opened: 'https://vibeoffice.work/updater/?from_version=v2026.08.16-000000&current_platform=web' });
  assert.deepEqual(shot('', [{ js: update }], '--pre', stub, '--qs', 'web&version=v2026.10.01-000000').r,
    { button: false, check: false, dialog: false, opened: '' }, 'two weeks old');
});

test('a version is outdated two calendar months after its day', () => {
  const { r } = shot('', [{ js: `const o = (v, y, m, d) => VO.isOutdated(v, new Date(y, m - 1, d)); return JSON.stringify([
    o('v2026.08.16-abc123', 2026, 10, 16), o('v2026.08.17-abc123', 2026, 10, 16), o('v2026.12.31-abc123', 2027, 3, 2),
    o('v2026.12.31-abc123', 2027, 3, 3), o('2026.08.8f8494', 2026, 10, 1), o('dev', 2099, 1, 1), VO.isOutdated()]);` }]);
  assert.deepEqual(r, [true, false, false, true, true, false, false], 'a dev build is never outdated');
});

test('a version\'s age in whole calendar months', () => {
  const { r } = shot('', [{ js: `const m = (v, y, mo, d) => VO.monthsOld(v, new Date(y, mo - 1, d)); return JSON.stringify([
    m('v2026.08.16-abc123', 2026, 10, 16), m('v2026.08.17-abc123', 2026, 10, 16), m('v2025.10.16-abc123', 2026, 10, 16),
    m('v2026.08.31-abc123', 2026, 10, 30), m('v2026.08.31-abc123', 2026, 10, 31), m('v2026.12.31-abc123', 2027, 3, 3),
    m('2026.08.8f8494', 2026, 10, 1), m('v2026.10.16-abc123', 2026, 10, 16), m('dev', 2099, 1, 1)]);` }]);
  assert.deepEqual(r, [2, 1, 12, 1, 2, 2, 2, 0, null], 'a dev build has no age');
});

test('versions compare by date alone', () => {
  const { r } = shot('', [{ js: `const c = VO.compareVersions; return JSON.stringify({
    older: c('v2026.10.10-ffffff', 'v2026.10.11-000000'), same: c('v2026.10.11-abc123', 'v2026.10.11-def456'),
    newer: c('v2026.11.01-abc123', 'v2026.10.31-abc123'), year: c('v2027.01.01-abc123', 'v2026.12.31-abc123'),
    legacy: c('2026.10.8f8494', 'v2026.10.01-abc123'), upper: c('V2026.10.11-abc123', 'v2026.10.11-def456'), devNewest: c('dev', 'v2099.12.31-abc123'), devDev: c('dev', 'dev'),
    web: VO.VERSION, label: document.querySelector('.version').parentNode.textContent.includes('Ver.') });` }]);
  assert.deepEqual(r, { older: -1, same: 0, newer: 1, year: 1, legacy: -1, upper: 0, devNewest: 1, devDev: 0, web: 'dev', label: false });
});

test('web: the host is the browser', () => {
  const { out, r } = shot('', [{ js: state }]);
  assert.equal(r.kind, 'web');
  assert.equal(r.platform, 'web');
  assert.equal(r.path, '/');
  assert.match(out, /no errors, no external requests/);
});
