/* VibeOffice suite services, shared by the start page and the apps (classic script, window.VO).
 *
 * - Which app opens which file (VO.APPS, VO.appFor).
 * - The host (VO.host): browser or desktop app, the files the system asks us to open.
 * - The service worker (public/sw.js), registered from every page.
 * - Files the start page hands to an app, and the Recent Files list. Both live in IndexedDB
 *   ('vibeoffice', stores 'handoff' and 'recent') in this browser only; nothing is uploaded.
 * - Unsaved versions: while a document has unsaved changes the app's own format is written (autosave)
 *   into its Recent Files entry as a draft; saving clears it. The Start Center, the apps' File menus and
 *   task panes and the Document Recovery pane all show the one list.
 *
 * An app page is opened as  <app>/?new           a blank document instead of the sample
 *                           <app>/?template=<id>  a new document from one of the app's templates
 *                           <app>/?open=<id>      a file handed over by VO.handoff (read once, then deleted)
 *                           <app>/?recent=<key>   a Recent Files entry (&draft: its unsaved version)
 * The app reads VO.launch while it starts, then calls VO.attach(app, { open, thumb }), and reports
 * VO.opened(file) after it opens a file and VO.saved(name, blob) after it saves in its own format.
 */
(function () {
  'use strict';
  const VO = (window.VO = window.VO || {});   // keep the language choice made before the first paint
  VO.VERSION = 'dev';   // the published copy says v<yyyy>.<mm>.<dd>-<6 hex> (tools/current-tree-version.sh, stamped by the sync script)
  if (window.__VO_HOST__ && window.__VO_HOST__.appVersion) VO.VERSION = window.__VO_HOST__.appVersion;   // the desktop app's build
  /** a version's date as a number (20261011), or null for a version without one (dev); copies before
   *  v<yyyy>.<mm>.<dd>-<hex> said <yyyy>.<mm>.<hex>, day 0 */
  const versionDay = (v) => {
    const m = /^v?(\d{4})\.(\d{2})(?:\.(\d{2})-[0-9a-f]{6})?/i.exec(v || '');
    return m ? +(m[1] + m[2] + (m[3] || '00')) : null;
  };
  /** compare two versions by their dates alone: < 0 if a is older than b, 0 for the same day, > 0 if newer;
   *  a version without a date (dev) is newer than any dated one */
  VO.compareVersions = (a, b) => {
    const x = versionDay(a), y = versionDay(b);
    return x === y ? 0 : x === null ? 1 : y === null ? -1 : Math.sign(x - y);
  };
  const BASE = new URL('..', document.currentScript.src).href;   // public/
  VO.base = BASE;

  /* ------------------------------------------------------------ apps */
  VO.APPS = {
    quire: { name: 'Quire', kind: 'Document', ext: ['docx', 'docm', 'dotx', 'dotm', 'doc', 'dot', 'rtf', 'htm', 'html', 'txt', 'md', 'markdown', 'zip'] },
    ledger: { name: 'Ledger', kind: 'Spreadsheet', ext: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xls', 'xlt', 'csv', 'tsv', 'prn', 'xml'] },
    lectern: { name: 'Lectern', kind: 'Presentation', ext: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm', 'ppt', 'pps', 'pot'] },
  };
  VO.ext = (name) => (/\.([^./\\]+)$/.exec(name || '') || ['', ''])[1].toLowerCase();
  /** the app for a file name, or null */
  VO.appFor = (name) => { const e = VO.ext(name); return Object.keys(VO.APPS).find((a) => VO.APPS[a].ext.includes(e)) || null; };
  VO.accept = () => Object.values(VO.APPS).flatMap((a) => a.ext.map((e) => '.' + e)).join(',');
  /** VO.url('quire', { new: true }) -> .../quire/?new */
  VO.url = (app, params) => {
    const qs = Object.keys(params || {}).map((k) => (params[k] === true ? k : k + '=' + encodeURIComponent(params[k]))).join('&');
    return new URL(app + '/', BASE).href + (qs ? '?' + qs : '');
  };

  /** the page's icon (browser tab) from an inline <svg> */
  VO.setIcon = (svg) => {
    if (!/xmlns=/.test(svg)) svg = svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
    let link = document.querySelector('link[rel="icon"]');
    if (!link) link = document.head.appendChild(Object.assign(document.createElement('link'), { rel: 'icon' }));
    link.type = 'image/svg+xml';
    link.href = 'data:image/svg+xml,' + encodeURIComponent(svg);
  };

  /* ------------------------------------------------------------ the host */
  /** VO.host: the one interface between the pages and what runs them (docs/desktop.md, Host interface).
   *  kind 'web': a browser tab or the installed PWA; 'desktop': the native app (desktop/), which sets
   *  window.__VO_HOST__ and answers the host_* commands (desktop/src/host.rs). Pages use VO.host, never
   *  Tauri's or the system's file APIs directly. */
  const native = window.__VO_HOST__ && window.__TAURI_INTERNALS__ ? window.__VO_HOST__ : null;
  /* the lite edition (desktop/src/lite.rs): one window, its start page (lite.html) and each app's page in turn; it
     keeps only the settings and the files in ~/VibeOffice/Documents: no Recent Files, previews or drafts */
  const lite = !!(native && native.edition === 'lite');
  /* for its small screen's styles (common/ui.css: no status bar but its zoom buttons) */
  if (lite) document.documentElement.classList.add('vo-lite');
  const call = (cmd, args, options) => window.__TAURI_INTERNALS__.invoke(cmd, args, options);
  /* an app page, not the Start Center (the desktop app: only the Start Center lists Recent Files) */
  const appPage = !!document.documentElement.dataset.app;
  /* the one-time word of a window reloaded for a file (?doc&t=): its document is given only to the page with it */
  const docWord = new URLSearchParams(location.search).get('t');
  /** bytes for a command's body (raw, never JSON) */
  const raw = async (blob) => new Uint8Array(await blob.arrayBuffer());
  VO.host = {
    kind: native ? 'desktop' : 'web',
    platform: native ? native.platform : 'web',
    version: native ? native.version : 0,
    /** whether this edition looks for its updates (updater/current.json): not the Mac App Store's, which the store
     *  updates and which has no network, nor the lite edition's, which its package updates */
    updates: !native || (native.edition !== 'appstore' && !lite),
    /** the lite edition's start page and window (desktop/src/lite.rs), else null: list() the documents in
     *  ~/VibeOffice/Documents, [{ name, app, size, modified }] newest first; open(name), create(page, start) a new
     *  document (start: { new } or { template }), home() and licenses() send the one window there; quit(); show(on): a
     *  slide show starts or ends, and with an HDMI screen the window goes there for it (true when it went) */
    lite: lite ? {
      list: () => call('host_lite_list'),
      open: (name) => call('host_lite_open', { name }),
      create: (page, start) => call('host_lite_new', { page, start }),
      home: () => call('host_lite_home'),
      licenses: () => call('host_lite_licenses'),
      quit: () => call('host_lite_quit'),
      show: (on) => call('host_lite_show', { on: !!on }),
    } : null,
    /** open a web or mail address in the system's browser: a new tab on the web, the default browser in the
     *  desktop app (which also sends every link that leaves the app there) */
    async openUrl(url) {
      if (native) return call('host_open_url', { url: String(url) });
      window.open(url, '_blank', 'noopener');
    },
    /** the whole UI's scale (1 = 100 %), or null where the page cannot set it (the web: the browser's own
     *  zoom, Ctrl + and Ctrl −, does it); the desktop app keeps it for the next start */
    /* the size of the whole UI in the desktop app (1 = 100 %): the setting the app zooms every window by */
    async uiScale() { return native ? +VO.settings.get('vibeoffice.uiScale') || 1 : null; },
    async setUiScale(scale) {
      if (!native) return;
      if (!(+scale >= 0.5 && +scale <= 4)) throw new RangeError('the UI size is from 0.5 to 4');
      VO.settings.set('vibeoffice.uiScale', String(+scale));
      await VO.settings.kept();
    },
    /** open one of the suite's pages in a window of its own: the desktop app makes a full window (host and
     *  all), told only the page (quire, ledger, lectern, licenses) and how it starts (VO.launch), never an
     *  address; the web opens a window or tab, false if the browser blocked it */
    async newWindow(url) {
      if (native) {
        const u = new URL(url, location.href), path = u.pathname.slice(new URL(BASE).pathname.length);
        const page = path === 'oss-licenses.html' ? 'licenses' : path.replace(/\/(index\.html)?$/, '');
        const start = {};
        for (const [k, v] of u.searchParams) start[k] = k === 'new' || k === 'draft' ? true : v;
        if (start.recent) await VO.host.openRecent(start.recent, start.draft);   // a Recent Files card
        else await call('host_new_window', { page, start });
        return true;
      }
      return !!window.open(url, '_blank');
    },
    /* The desktop app: a window holds one document, and only the native side gives a window its file
       (desktop/src/doc.rs). A page names no file: it asks for the system's Open dialog, a Recent Files entry
       by its key, or a new window for bytes it has. */
    /** the system's Open dialog: { opened, here }. In an app window the first file of its app is to take the
     *  window's place (here: its name; reopen() then reloads the window with it); every other file opens in a window
     *  of its own, or comes forward in the one that has it (opened: how many) */
    async open() {
      const labels = { title: VO.t('Open'), all: VO.t('All files VibeOffice opens'), quire: VO.t('Documents'), ledger: VO.t('Spreadsheets'), lectern: VO.t('Presentations') };
      return call('host_open', { labels });
    },
    /** the file open() picked takes this window's place: the window reloads with it (go: false lets it go) */
    async reopen(go) { return call('host_reopen', { go: !!go }); },
    /** the Start Center's Recent Files entry (draft: its unsaved version) in a window of its own, or the window that
     *  has it brought forward */
    async openRecent(key, draft) { return call('host_open_recent', { key, draft: !!draft }); },
    /** an app window's Recent Files: the Start Center, the one window that lists them, on its Recent Files (the lite
     *  edition: its start page's list of documents, once this page's document has gone) */
    async showRecent() { return lite ? VO.openDialog() : call('host_show_recent'); },
    /** a document as bytes in a new window: a file dropped on a page, or one a page made (made: never saved) */
    async newDocument(blob, name, made) {
      const headers = { 'X-Name': encodeURIComponent(name) };
      if (made) headers['X-Made'] = '1';
      return call('host_new_document', await raw(blob), { headers });
    },
    /** this window's document: { info: { fileName, name, saved, draft, lossState, bound }, file: File }, or null
     *  (blank, a template, the sample) */
    async document() {
      const info = native && await call('host_document', { t: docWord });
      if (!info) return null;
      return { info, file: new File([await call('host_document_bytes', { t: docWord })], info.fileName) };
    },
    /** the Window menu: this app's windows [{ id, title, current }], and one brought forward */
    async windows() { return native ? call('host_windows') : []; },
    async focus(id) { return call('host_focus', { id }); },
    /* Saving (common/core.js, L.file): the window's own file, its name, whether it may be changed */
    file: { bound: false, readOnly: false, name: '' },
    /** Save: bytes into the window's own file (replace: even though another program changed it) */
    async save(bytes, replace) { return call('host_save', bytes, replace ? { headers: { 'X-Replace': '1' } } : undefined); },
    /** the Save dialog ({ purpose, types: [{ id, label }], type, name, title }): { name, type } or null */
    async choose(args) { return call('host_choose', args); },
    /** bytes into what the Save dialog picked: { name, type, bound } */
    async writeChosen(bytes) { return call('host_write_chosen', bytes); },
    async cancelChoice() { return call('host_cancel_choice'); },
    /** in the Mac App Sandbox (both editions), a save it refused may ask for the file's folder: allowAccess (how:
     *  'save' or 'chosen', text: the folder chooser's message) is true once the user let the app in */
    sandboxed: !!(native && native.sandboxed),
    async allowAccess(how, text) { return call('host_allow_access', { how, text }); },
    /** File ▸ Versions...: the versions of this window's file its saves replaced [{ id, time, size }], and one
     *  opened in a new window */
    async versions() { return native ? call('host_versions') : []; },
    async openVersion(id) { return call('host_open_version', { id }); },
    /** close this window if it is an app window the suite opened; false for the Start Center's (and on the web) */
    async closeWindow() { return native && !lite ? call('host_close_window') : false; },
    /** fn(files) for the files the system asks us to open: the installed PWA's file handlers (Open with, a
     *  double-click on an associated file); the page that calls this takes them. The desktop app opens those
     *  itself, each in a window of its own */
    onFiles(fn) {
      if (!native && 'launchQueue' in window) {
        window.launchQueue.setConsumer(async (params) => {
          if (params.files && params.files.length) fn(await Promise.all(params.files.map((h) => h.getFile())));
        });
      }
    },
  };

  /** an app's name with this edition's, for Help ▸ About and Getting Started, in English in every language: Quire Web Edition, Quire
   *  Windows Edition, Quire macOS Edition (DMG) or (App Store), Quire Linux Edition */
  VO.editionName = (app) => {
    if (!native) return app + ' Web Edition';
    if (lite) return app + ' Lite Edition';
    const system = { windows: 'Windows', macos: 'macOS', linux: 'Linux' }[native.platform] || native.platform;
    return `${app} ${system} Edition` + (native.platform === 'macos' ? (native.edition === 'appstore' ? ' (App Store)' : ' (DMG)') : '');
  };

  /* ------------------------------------------------------------ updates */
  /** updater/current.json on the published site says, per policy and platform, the version an edition must
   *  be at least: older than "show-update-dialog" asks with a dialog, older than "show-update-button" shows an
   *  Update button. Versions compare by date (VO.compareVersions); a dev build never asks. The desktop app
   *  checks when the Start Center opens; nothing about the user or their files is sent. */
  VO.UPDATES = (native && native.updates) || 'https://vibeoffice.work/updater/current.json';
  /** updater/whatsnew.json: { "<version>": "<plain text>" }, one list for every platform */
  VO.WHATSNEW = (native && native.whatsnew) || 'https://vibeoffice.work/updater/whatsnew.json';
  /** the page the Update button opens (in the system's browser): it gets this edition's version and platform, the
   *  Mac App Store's as macos_store (the page then sends people to the App Store) */
  VO.updateUrl = () => 'https://vibeoffice.work/updater/?from_version=' + encodeURIComponent(VO.VERSION)
    + '&current_platform=' + encodeURIComponent(native && native.edition === 'appstore' ? 'macos_store' : VO.host.platform);
  /** how many whole calendar months a version (by default this one) is old on `now`, by its date; null for a dev build */
  VO.monthsOld = (v = VO.VERSION, now = new Date()) => {
    const d = versionDay(v);
    if (d === null) return null;
    const y = Math.floor(d / 10000), m = Math.floor(d / 100) % 100 - 1, day = d % 100 || 1;
    let n = (now.getFullYear() - y) * 12 + now.getMonth() - m;
    while (n > 0 && now < new Date(y, m + n, day)) n--;   // a 31st's mark in a shorter month runs on: Feb 31 is Mar 3
    return n;
  };
  /** whether a version (by default this one) is two months old or more on `now`: the Start Center then offers to
   *  check for updates, wherever the version runs (the site, the installed PWA's cache, the desktop app, the Mac App
   *  Store's too) and whether or not current.json can be read. A dev build never is. */
  VO.isOutdated = (v = VO.VERSION, now = new Date()) => VO.monthsOld(v, now) >= 2;
  /** { ask: 'dialog', 'button' or null (up to date), news: [{ version, text }] }: when an update is due, the
   *  whatsnew.json entries newer than this version, newest first (none if that file cannot be read).
   *  Throws when current.json cannot be read (offline). */
  VO.checkUpdate = async () => {
    const res = await fetch(VO.UPDATES, { cache: 'no-store' });
    if (!res.ok) throw new Error('updates: HTTP ' + res.status);
    const policy = (await res.json()) || {};
    const older = (k) => policy[k] && typeof policy[k][VO.host.platform] === 'string'
      && VO.compareVersions(VO.VERSION, policy[k][VO.host.platform]) < 0;
    const ask = older('show-update-dialog') ? 'dialog' : older('show-update-button') ? 'button' : null;
    let all = {};
    if (ask) try { const r = await fetch(VO.WHATSNEW, { cache: 'no-store' }); if (r.ok) all = (await r.json()) || {}; } catch (e) { /* the dialog without notes */ }
    const news = Object.entries(all)
      .filter(([v, text]) => typeof text === 'string' && versionDay(v) !== null && VO.compareVersions(VO.VERSION, v) < 0)
      .sort(([a], [b]) => VO.compareVersions(b, a))
      .map(([version, text]) => ({ version, text }));
    return { ask, news };
  };

  /* ------------------------------------------------------------ open source licenses */
  /** About ▸ Open source licenses: the credits page (made by tools/oss-licenses.py) in a window or tab of its own */
  VO.LICENSES = BASE + 'oss-licenses.html';
  VO.showLicenses = () => VO.host.newWindow(VO.LICENSES).catch((e) => console.error('licenses:', e));

  /* ------------------------------------------------------------ fonts on this computer */
  /** the font families Local Font Access listed (Chrome and Edge, after the user allows it), kept in the
   *  settings for every app of the suite; the Font boxes show them (L.localFonts in core.js) */
  const LOCAL_FONTS = 'vibeoffice.localFonts';
  let localFonts = null;
  VO.settings.on((key) => { if (key === LOCAL_FONTS || key === null) localFonts = null; });
  VO.localFonts = {
    supported: () => typeof window.queryLocalFonts === 'function',
    list() {
      if (!localFonts) { try { localFonts = JSON.parse(VO.settings.get(LOCAL_FONTS)) || []; } catch (e) { localFonts = []; } }
      return localFonts;
    },
    /** ask the browser for the list (it asks the user the first time) and keep it; the number of families */
    async load() {
      const families = Array.from(new Set((await window.queryLocalFonts()).map((f) => f.family))).sort((x, y) => x.localeCompare(y));
      localFonts = families;
      VO.settings.set(LOCAL_FONTS, JSON.stringify(families));   // not kept: for this page only
      return families.length;
    },
    forget() { localFonts = []; VO.settings.remove(LOCAL_FONTS); },
  };

  /* ------------------------------------------------------------ service worker */
  /* The published site, https://vibeoffice.work/, is cached for offline use, one version per page (sw.js); the
     worker is registered there and on 127.0.0.1 (any port: a published copy tried locally; the dev server's own
     files are unstamped, so the worker caches nothing there), and anywhere else (preview deployments) any earlier
     registration is removed. While the worker caches a version it reports to every page: 'vo-cache' on window, detail { state:
     'progress', 'done' or 'failed', update (an earlier version was cached: this is an update), progress (0 to 1),
     version }. updateViaCache 'none': the browser always asks the server for sw.js, so a new version is found on
     the next visit. */
  if (VO.host.kind === 'web' && 'serviceWorker' in navigator) {
    if (location.origin === 'https://vibeoffice.work' || location.hostname === '127.0.0.1') {
      navigator.serviceWorker.register(BASE + 'sw.js', { updateViaCache: 'none' }).catch((e) => console.warn('service worker:', e.message));
      navigator.serviceWorker.addEventListener('message', (e) => {
        if (e.data && e.data.type === 'vo-cache') window.dispatchEvent(new CustomEvent('vo-cache', { detail: e.data }));
      });
      navigator.serviceWorker.startMessages();
    } else {
      navigator.serviceWorker.getRegistrations().then((all) => all.forEach((r) => r.unregister())).catch(() => {});
    }
  }

  /* Installing is offered only on the published site and on 127.0.0.1 (trying it locally): anywhere else (preview
     deployments, a copy) the browser's own install prompt is cancelled, and the Start Center shows no Install button
     (start.js) */
  VO.offersInstall = location.origin === 'https://vibeoffice.work' || location.hostname === '127.0.0.1';
  if (!VO.offersInstall) window.addEventListener('beforeinstallprompt', (e) => e.preventDefault());

  /* ------------------------------------------------------------ IndexedDB */
  let dbp = null;
  function db() {
    return dbp || (dbp = new Promise((resolve, reject) => {
      const r = indexedDB.open('vibeoffice', 1);
      r.onupgradeneeded = () => {
        r.result.createObjectStore('recent', { keyPath: 'key' });
        r.result.createObjectStore('handoff', { keyPath: 'id' });
      };
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    }));
  }
  /** run fn(store) in one transaction; resolves with the result of the request fn returns, once committed */
  async function tx(name, mode, fn) {
    const t = (await db()).transaction(name, mode);
    return new Promise((resolve, reject) => {
      let out;
      const req = fn(t.objectStore(name));
      if (req) req.onsuccess = () => { out = req.result; };
      t.oncomplete = () => resolve(out);
      t.onerror = t.onabort = () => reject(t.error);
    });
  }
  const chan = !native && 'BroadcastChannel' in window ? new BroadcastChannel('vibeoffice') : null;   // the desktop app: vo-recent
  const listeners = new Set();
  const notify = () => { VO.recent.refresh(); for (const fn of listeners) if (!fn.cacheOnly) fn(); };   // cache users redraw after the refresh
  const changed = () => { if (chan) chan.postMessage('recent'); notify(); };
  if (chan) chan.addEventListener('message', (e) => { if (e.data === 'recent') notify(); });
  /** call fn when the Recent Files list changes (here or in another page); returns the function that stops it */
  VO.onRecentChange = (fn) => { listeners.add(fn); return () => listeners.delete(fn); };

  /* ------------------------------------------------------------ Looks */
  /* the chrome's colours (common/looks.css) and, for a look with layout 'ribbon', the ribbon instead of menus and
     toolbars (common/ribbon.js); one choice for the whole suite: <html data-look data-layout> is set before
     the first paint by each page's head script (common/look.js) from the setting 'vibeoffice.look'; changing it
     here repaints every open window of the suite. Canvas-drawn chrome (rulers, sheet headers) redraws on 'vo-look'. */
  const N = (s) => s;   // UI text, translated where it is shown (VO.N, common/i18n.js)
  N('Blank Document'); N('Blank Workbook'); N('Blank Presentation');   // templates.json's Blank entries (tools/templates.mjs)
  const LOOKS = [{ id: 'classic', name: N('Classic'), note: 'The classic blue toolbars and menus.' },
    { id: 'paper', name: N('Paper'), note: 'Light and calm: warm paper tones, one teal accent, rounded shapes.' },
    { id: 'paper2016', name: N('Paper 2016'), layout: 'ribbon', note: 'Paper with tabs of commands, grouped by task, instead of menus and toolbars.' }];
  const LOOK = 'vibeoffice.look';
  const storedLook = () => VO.settings.get(LOOK) || 'classic';
  const applyLook = (id) => {
    const root = document.documentElement;
    if (!LOOKS.some((l) => l.id === id)) return;
    const changed = root.dataset.look !== id;
    root.dataset.look = id;
    root.dataset.layout = LOOKS.find((l) => l.id === id).layout || 'bars';
    const meta = document.querySelector('meta[name="theme-color"]');
    const c = getComputedStyle(root).getPropertyValue('--theme-color').trim();
    if (meta && c) meta.content = c;
    if (changed) window.dispatchEvent(new Event('vo-look'));
  };
  VO.look = {
    /** [{ id, name, note, layout }] */
    list: () => LOOKS.map((l) => ({ ...l })),
    get: () => document.documentElement.dataset.look || 'classic',
    /** 'bars' (menus and toolbars) or 'ribbon' */
    layout: () => document.documentElement.dataset.layout || 'bars',
    set(id) {
      if (!LOOKS.some((l) => l.id === id)) return;
      VO.settings.set(LOOK, id);   // not kept: this window still changes
      applyLook(id);
    },
  };
  applyLook(storedLook());   // a page without the head script, or a stored look that changed meanwhile
  VO.settings.on((key) => { if (key === LOOK || key === null) applyLook(storedLook()); });

  /* ------------------------------------------------------------ Recent Files */
  const MAX_RECENT = 30;
  const MAX_BYTES = 50 * 1024 * 1024;   // bigger files are not kept (open them from disk again)
  const prune = async () => {
    /* entries with unsaved changes are never dropped */
    const all = (await VO.recent.list()).filter((e) => !e.draft);
    if (all.length > MAX_RECENT) await tx('recent', 'readwrite', (s) => { for (const e of all.slice(MAX_RECENT)) s.delete(e.key); });
  };
  /* Safari clears a site's data after about 7 days without a visit, unsaved versions included; once a
     document has unsaved changes, ask the browser to keep the data (Safari and Chrome decide without
     asking the user; Firefox may ask) */
  let persistAsked = false;
  const keepData = () => {
    if (persistAsked || !navigator.storage || !navigator.storage.persist) return;
    persistAsked = true;
    navigator.storage.persisted().then((yes) => yes || navigator.storage.persist()).catch(() => {});
  };
  VO.recent = {
    /** newest first: { key, app, name, size, time, saved (it has a file), file: Blob|null (opened or saved),
     *  thumb: Blob|null, draft: { file: Blob, name, time } | null (unsaved changes) }; a never-saved document has
     *  only a draft. The desktop app's entries have no file or thumb, and say open (in a window now) and here (in
     *  this one) */
    async list() {
      if (native) return call('host_recent');
      const all = (await tx('recent', 'readonly', (s) => s.getAll())) || [];
      for (const e of all) e.saved = !!e.file;
      return all.sort((a, b) => b.time - a.time);
    },
    get: (key) => tx('recent', 'readonly', (s) => s.get(key)),
    /** an entry's picture (a Blob), or null */
    async preview(e) {
      if (!native) return e.thumb || null;
      return e.preview ? new Blob([await call('host_recent_preview', { key: e.key })], { type: 'image/jpeg' }) : null;
    },
    /** the list as last read (for menus, which are built synchronously) */
    cache: [],
    refresh: (() => { let t = 0; return () => { clearTimeout(t); t = setTimeout(() => VO.recent.list().then((l) => { VO.recent.cache = l; for (const fn of listeners) if (fn.cacheOnly) fn(); }, () => {}), 30); }; })(),
    /** this app's entries, newest first, from the cache */
    forApp: (app) => VO.recent.cache.filter((e) => e.app === (app || (current && current.app))),
    /** add or refresh the entry for this file name (a newer copy replaces the older one, its thumbnail stays);
     *  keepDraft: opening a file keeps unsaved changes made to it, saving it clears them */
    async put(app, name, blob, keepDraft) {
      if (!app || !blob || blob.size > MAX_BYTES) return null;
      const key = app + ':' + name;
      const old = await VO.recent.get(key);
      await tx('recent', 'readwrite', (s) => s.put({ key, app, name, size: blob.size, time: Date.now(), file: blob, thumb: old ? old.thumb : null, draft: keepDraft && old ? old.draft || null : null }));
      await prune();
      changed();
      return key;
    },
    /** store the unsaved version of a document */
    async setDraft(key, app, name, draftName, blob, lossState) {
      const e = (await VO.recent.get(key)) || { key, app, name, size: 0, file: null, thumb: null };
      if (!e.file) e.name = name;
      e.draft = { file: blob, name: draftName || name, time: Date.now(), lossState };
      e.time = e.draft.time;
      await tx('recent', 'readwrite', (s) => s.put(e));
      changed();
      keepData();
    },
    /** throw the unsaved version away; a document that was never saved leaves the list */
    async discardDraft(key) {
      if (native) return call('host_discard_draft', { key });
      const e = await VO.recent.get(key);
      if (!e) return;
      if (!e.file) await tx('recent', 'readwrite', (s) => s.delete(key));
      else { e.draft = null; await tx('recent', 'readwrite', (s) => s.put(e)); }
      changed();
    },
    async setThumb(key, thumb) {
      const e = await VO.recent.get(key);
      if (!e) return;
      e.thumb = thumb;
      await tx('recent', 'readwrite', (s) => s.put(e));
      changed();
    },
    async remove(key) { if (native) return call('host_recent_remove', { key }); await tx('recent', 'readwrite', (s) => s.delete(key)); changed(); },
    async clear() { if (native) return call('host_recent_clear'); await tx('recent', 'readwrite', (s) => s.clear()); changed(); },
  };
  /* the desktop app keeps Recent Files itself (~/.vibeoffice, desktop/src/store.rs) and says when they change; a
     window's own entry changes only through its document (the draft, the preview), never by a key */
  if (native && !appPage) window.addEventListener('vo-recent', notify);
  /** whether an entry's document is open (in the desktop app: in any window; on the web: in this page) */
  const isOpen = (e) => (native ? e.open : openKeys.has(e.key));

  /* ------------------------------------------------------------ handing files to apps */
  /** where an app page opens: a new window in the desktop app or when installed (the start page stays, as in
   *  LibreOffice), else this tab. The desktop app's windows each keep their one page, so there a window that
   *  can't be opened is an error for the caller to show, not a page change */
  VO.go = async (url) => {
    if (lite) return liteGo(url);
    if (VO.host.kind === 'desktop') return VO.host.newWindow(url);
    const standalone = matchMedia('(display-mode: standalone)').matches || matchMedia('(display-mode: window-controls-overlay)').matches;
    try { if (standalone && await VO.host.newWindow(url)) return; } catch (e) { console.error('new window:', e); }
    location.href = url;
  };
  /** File ▸ Exit, after the app has closed its documents: an app window the Start Center opened
   * (installed) closes and shows the Start Center behind it; a tab goes to the Start Center */
  VO.exit = async () => {
    if (lite) { leaving = true; return VO.host.lite.home().catch((e) => { leaving = false; say(e); }); }
    if (await VO.host.closeWindow().catch(() => false)) return;   // the desktop app's app windows
    if (window.opener && !window.opener.closed) { window.close(); if (window.closed) return; }
    location.href = BASE;
  };
  /* the desktop app: an app window's close button asks its page (vo-close), which runs its "Save changes?" question
     (hooks.close) and closes the window itself; Cancel keeps it open. A page that never takes it in hand gets the
     native side's own question after 10 s (desktop/src/doc.rs) */
  let closing = false;
  if (native) window.addEventListener('vo-close', async () => {
    if (!current) { call('host_close_window').catch(() => {}); return; }   // no document here
    call('host_closing').catch(() => {});
    if (closing) return;
    closing = true;
    try { if (!current.hooks.close || await current.hooks.close()) await call('host_close_window'); } catch (e) { console.error('close:', e); } finally { closing = false; }
  });
  /** open a file in the app made for it; false if no app opens this kind of file */
  VO.handoff = async (file) => {
    const app = VO.appFor(file.name);
    if (!app) return false;
    if (native) { await VO.host.newDocument(file, file.name); return true; }   // a window for its bytes
    const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    const day = Date.now() - 86400000;
    await tx('handoff', 'readwrite', (s) => {
      s.put({ id, name: file.name, file, time: Date.now() });
      s.openCursor().onsuccess = (e) => { const c = e.target.result; if (!c) return; if (c.value.time < day) c.delete(); c.continue(); };
    });
    await VO.go(VO.url(app, { open: id }));
    return true;
  };

  /* ------------------------------------------------------------ the app side */
  const q = new URLSearchParams(location.search);
  /** how this app page was opened: null (plain: show the sample), { new }, { template: id }, { open: id } or { recent: key };
   *  the desktop app's { doc }: the window has a document (VO.host.document); n: the new document's number */
  VO.launch = q.has('open') ? { open: q.get('open') } : q.has('recent') ? { recent: q.get('recent'), draft: q.has('draft') }
    : q.has('doc') && native ? { doc: true } : q.has('template') ? { template: q.get('template') } : q.has('new') ? { new: true } : null;
  /* the desktop app numbers new documents across windows (Document2 in the second window that makes one): n */
  if (VO.launch && (VO.launch.new || VO.launch.template) && +q.get('n') > 0) VO.launch.n = +q.get('n');

  let current = null;   // { app, hooks }
  /** an app is ready: open the file it was launched with, if any, then show the Document Recovery pane when
   *  other documents of this app have unsaved changes.
   *  hooks: open(file, { draft, name, saved }), thumb() -> a canvas (or image) of the first page, message(text),
   *  templates() -> [{ id, name }] (read by tools/templates.mjs); for autosave: snapshot(doc) -> Blob in the
   *  app's format, docInfo(doc) -> { name, draftName }, isDirty(doc), current() -> doc, autosave() -> false
   *  when turned off; showRecovery() opens the Document Recovery pane. The desktop app's one document per window:
   *  pristine() -> whether the document is an untouched blank (a file opened goes into this window, else into a
   *  new one), close() -> false when the user cancels closing (the "Save changes?" question) */
  VO.attach = async (app, hooks) => {
    current = { app, hooks: hooks || {} };
    if (!native) VO.recent.refresh();
    readWindows();
    const l = VO.launch;
    if (l && !l.new) {
      if (history.replaceState) history.replaceState(null, '', location.pathname);   // a reload starts clean
      if (l.doc) await takeDocument();
      else if (l.open) {
        let rec = null;
        try {
          rec = await tx('handoff', 'readonly', (s) => s.get(l.open));
          if (rec) await tx('handoff', 'readwrite', (s) => s.delete(l.open));
        } catch (e) { console.error(e); }
        if (rec && rec.file) await current.hooks.open(asFile(rec.file, rec.name), {});
        else (current.hooks.message || alert)(VO.t('That file is no longer available here. Open it again from your computer.'));
      } else if (l.recent) await VO.openRecent(l.recent, { draft: l.draft });
    }
    /* (the desktop app: the Start Center shows them) */
    if (!native) setTimeout(async () => {
      const others = (await VO.recent.list()).filter((e) => e.app === app && e.draft && !isOpen(e));
      if (others.length && current.hooks.showRecovery) current.hooks.showRecovery();
    }, 700);
  };
  const asFile = (blob, name) => (blob instanceof File && blob.name === name ? blob : new File([blob], name, { type: blob.type }));
  const say = (e) => (current && current.hooks.message || alert)(String(e && e.message || e));
  /** the desktop app: open the document the native side gave this window as it was made */
  async function takeDocument() {
    let d = null;
    try { d = await VO.host.document(); } catch (e) { say(e); return; }
    if (!d) return;
    const i = d.info;
    VO.host.file = { bound: i.bound, readOnly: i.readOnly, name: i.name };
    await current.hooks.open(d.file, { draft: i.draft, name: i.name, saved: i.saved, lossState: i.lossState || undefined, made: !i.saved && !i.draft });
  }
  const pristine = () => !!(current && current.hooks.pristine && current.hooks.pristine());
  /** the desktop app: a window whose document is an untouched blank closes once the file asked for has opened in
   *  a window of its own, as Word's blank Document1 gives way (the page only closes itself: its claim gets it nothing) */
  /** File ▸ Open in the desktop app: the system's Open dialog. The first file of this app takes this window's place
   *  once its document has gone, saved or given up as when the window closes ("Save changes?"; Cancel keeps it):
   *  the window reloads with the file, as a fresh page. Other files open in windows of their own */
  VO.openDialog = async () => {
    if (lite) return leave(() => VO.host.lite.home());   // its list of documents is the Open dialog
    try {
      const r = await VO.host.open();
      if (!r || !r.here) return;
      const go = !current || !current.hooks.close || await current.hooks.close();
      leaving = go;
      await VO.host.reopen(go);
    } catch (e) { leaving = false; say(e); }
  };
  /* its document let go, the page leaves without the apps' "Leave the page?" (their beforeunload, added after this) */
  let leaving = false;
  /** the lite edition: the one window goes elsewhere (`to`, a host call) once this page's document has gone, saved or
   *  given up as when a window closes ("Save changes?"; Cancel keeps it); false when it stays */
  async function leave(to) {
    try {
      if (current && current.hooks.close && !(await current.hooks.close())) return false;
      leaving = true;
      await to();
      return true;
    } catch (e) { leaving = false; say(e); return false; }
  }
  /** VO.go in the lite edition: a new document of an app (?new, ?template=) or the licences page, in this window */
  function liteGo(url) {
    const u = new URL(url, location.href), path = u.pathname.slice(new URL(BASE).pathname.length);
    if (path === 'oss-licenses.html') return leave(() => VO.host.lite.licenses());
    const page = path.replace(/\/(index\.html)?$/, '');
    const start = u.searchParams.has('template') ? { template: u.searchParams.get('template') } : { new: true };
    return leave(() => VO.host.lite.create(page, start));
  }
  window.addEventListener('beforeunload', (e) => { if (leaving) e.stopImmediatePropagation(); }, true);
  /** a dialog of the page is open: the Mac's menu bar (vo-menu) waits for it, as a Mac app's menus do */
  VO.busyDialog = () => !!document.querySelector('.dlg-overlay, [aria-modal="true"]');
  /* the Mac's menu bar (desktop/src/main.rs, mac_menu_event): its File ▸ Open… and Help in the app window in front run
     the app's own Open and Help (the Start Center answers Open itself, start.js) */
  if (native) window.addEventListener('vo-menu', (e) => {
    const ui = window.L && window.L.ui;
    if (!current || !ui || VO.busyDialog() || !['open', 'help'].includes(e.detail)) return;
    ui.exec(e.detail);
  });
  /** whether a window holds one document (the desktop app, docs/desktop.md): the apps' New, Open and Close then
   *  make or close windows, and their Window menus list windows (VO.windowItems) */
  VO.oneDocument = !!native;
  /** the desktop app keeps one document per window: a file dropped on (or handed to) a window whose document isn't
   *  an untouched blank opens in a new window instead; true when it went there */
  VO.elsewhere = async (file) => {
    if (!native || !current || pristine()) return false;
    await VO.host.newDocument(file, file.name);
    return true;
  };
  /** the Window menu in the desktop app: this app's windows ({ label, run } items for L.ui menus), as last read;
   *  the list is read again whenever this window comes forward */
  let wins = [];
  const readWindows = () => { if (current) VO.host.windows().then((l) => { wins = l; }, () => {}); };
  if (native) window.addEventListener('focus', readWindows);
  VO.windowItems = () => wins.map((w, i) => {
    const title = (w.current ? document.title : w.title).replace(' - ' + VO.APPS[current.app].name, '');   // the document's name
    return { label: '&' + (i + 1) + ' ' + title.replace(/&/g, '&&'), raw: true, checked: () => w.current, run: () => VO.host.focus(w.id).catch(say) };
  });
  /** a document a page made (mail merge's letters) in a window of its own in the desktop app, as bytes in the
   *  app's format; false on the web, where the page keeps it */
  VO.made = async (blob, name) => {
    if (!native) return false;
    await VO.host.newDocument(blob, name, true);
    return true;
  };
  /** open a Recent Files entry in this app: its unsaved version when draft is true, else the saved file (the
   *  desktop app: the Start Center instead, the one window that lists them) */
  VO.openRecent = async (key, o) => {
    if (native) { VO.host.showRecent().catch(say); return; }   // the Start Center lists them
    let e = null;
    try { e = await VO.recent.get(key); } catch (err) { console.error(err); }
    const draft = !!(o && o.draft && e && e.draft);
    const blob = e && (draft ? e.draft.file : e.file);
    if (!blob) { (current.hooks.message || alert)(VO.t('That file is no longer available here. Open it again from your computer.')); return; }
    const f = asFile(blob, draft ? e.draft.name : e.name);
    if (draft) f.voDraft = key;
    await current.hooks.open(f, { draft, name: e.name, saved: !!e.file, lossState: draft ? e.draft.lossState : undefined });
  };
  /** File ▸ New… inside an app: Blank and the app's templates with their previews (the Start Center's
   *  Create File gallery, common/templates.json), in the app's own dialog. create(id): id '' is Blank. */
  VO.newDialog = async (app, create) => {
    const ui = window.L && window.L.ui;
    /* the desktop app: a new document in a window of its own */
    const pick = (id) => (native ? VO.go(VO.url(app, id ? { template: id } : { new: true })).catch(say) : create(id));
    if (!ui || !ui.dialog) return pick('');
    let list = null;
    try { list = (await (await fetch(BASE + 'common/templates.json')).json())[app]; } catch (e) { console.warn('templates:', e); }
    if (!list || !list.length) list = [{ id: '', name: 'Blank', img: null }];
    if (!document.getElementById('vo-gallery-css')) {
      const st = document.createElement('style');
      st.id = 'vo-gallery-css';
      st.textContent = '.vo-gal{display:grid;grid-template-columns:repeat(auto-fill,minmax(132px,1fr));gap:8px;max-height:min(62vh,540px);overflow:auto;padding:6px;background:var(--surface);border:1px solid var(--field-border)}'
        + '.vo-tpl{display:flex;flex-direction:column;gap:4px;padding:5px;border:1px solid transparent;border-radius:2px;background:none;text-align:left;cursor:default;font:inherit;color:var(--ink)}'
        + '.vo-tpl:hover,.vo-tpl:focus-visible{background:var(--menu-hot);border-color:var(--card-hot-border);outline:none}'
        + '.vo-tpl span{display:flex;align-items:center;justify-content:center;height:112px;background:var(--card-thumb-bg);border:1px solid var(--card-thumb-line);overflow:hidden}'
        + '.vo-tpl img{max-width:100%;max-height:100%;box-shadow:0 1px 2px var(--tip-shadow-c)}'
        + '.vo-tpl b{font-weight:normal;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
        + '.vo-gal-note{margin:0 0 6px}';
      document.head.appendChild(st);
    }
    const gal = document.createElement('div');
    gal.className = 'vo-gal';
    let d = null;
    for (const t of list) {
      const b = document.createElement('button');
      b.type = 'button'; b.className = 'vo-tpl'; b.title = VO.t(t.name);
      const pic = document.createElement('span');
      if (t.img) { const img = document.createElement('img'); img.src = BASE + t.img; img.alt = ''; pic.append(img); }
      const name = document.createElement('b');
      name.textContent = VO.t(t.name);
      b.append(pic, name);
      b.addEventListener('click', () => { if (d) d.close(-1); pick(t.id); });
      gal.append(b);
    }
    const body = document.createElement('div');
    const note = document.createElement('p');
    note.className = 'vo-gal-note';
    note.textContent = VO.t('Choose Blank or a template:');
    body.append(note, gal);
    d = ui.dialog({ title: VO.t({ quire: 'New Document', ledger: 'New Spreadsheet', lectern: 'New Presentation' }[app]), width: 700, body, buttons: [{ label: 'Cancel', cancel: true }], focus: gal.firstChild });
  };
  /** the open app's templates: [{ id, name }] */
  VO.templates = () => (current && current.hooks.templates ? current.hooks.templates() : []);
  /** a JPEG (at most 320 px) of the open document's first page, or null */
  VO.thumbnail = async () => {
    const make = current && current.hooks.thumb;
    const src = make && (await make());
    if (!src) return null;
    const k = Math.min(1, 320 / src.width, 320 / src.height);
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(src.width * k)); c.height = Math.max(1, Math.round(src.height * k));
    const g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
    g.drawImage(src, 0, 0, c.width, c.height);
    return new Promise((r) => c.toBlob(r, 'image/jpeg', 0.88));
  };
  /** the thumbnail for a Recent Files entry, a little after the document has been drawn (the desktop app: this
   *  window's entry, if it has one) */
  function thumbLater(key) {
    if (lite || (!key && !native)) return;
    setTimeout(async () => {
      try {
        const blob = await VO.thumbnail();
        if (blob && native) await call('host_set_preview', await raw(blob));
        else if (blob) await VO.recent.setThumb(key, blob);
      } catch (e) { console.warn('thumbnail:', e); }
    }, 600);
  }
  /* ------------------------------------------------------------ documents, autosave */
  const docKeys = new WeakMap();   // the app's document object -> its Recent Files key
  const openKeys = new Set();      // keys of the documents open in this page (left out of Document Recovery)
  /* a document's unsaved version is written 15 s after its last change, at most once every 2 minutes, larger
     documents less often (12 s a MB: 50 MB, every 10 minutes, as Word's AutoRecover), slow ones too (30 × the time
     the last write took); typing on doesn't put it off past that gap */
  const drafts = new WeakMap();   // doc -> { timer, since: its first change not yet written, last: the last write, gap, big }
  const DRAFT_IDLE = 15000, DRAFT_GAP = 120000, DRAFT_PER_MB = 12000;
  const draftOf = (doc) => { let s = drafts.get(doc); if (!s) drafts.set(doc, (s = { timer: 0, since: null, last: -Infinity, gap: DRAFT_GAP, big: false })); return s; };
  const bind = (doc, key) => { if (doc) docKeys.set(doc, key); openKeys.add(key); };
  /** the app opened a file (from any source) as doc */
  VO.opened = (file, doc) => {
    if (!current) return;
    if (native) { thumbLater(); return; }   // the native side made its entry as it gave the window the file
    if (file.voDraft) { bind(doc, file.voDraft); return; }   // an unsaved version: keep its entry as it is
    VO.recent.put(current.app, file.name, file, true).then((key) => { if (key) bind(doc, key); thumbLater(key); }, (e) => console.warn('recent files:', e));
  };
  /** the app saved doc in its own format as name: that is its entry now, without unsaved changes */
  VO.saved = (name, blob, doc) => {
    if (!current) return;
    const old = doc && docKeys.get(doc);
    if (doc) { const s = draftOf(doc); clearTimeout(s.timer); s.since = null; s.last = Date.now(); }   // a save is a write too
    if (doc && current.hooks.isDirty && current.hooks.isDirty(doc)) VO.changed(doc);   // changed while it was saved
    if (native) { thumbLater(); return; }   // the native side cleared the draft as it wrote the file
    VO.recent.put(current.app, name, blob, false).then(async (key) => {
      if (!key) return;
      if (old && old !== key) { await VO.recent.discardDraft(old); openKeys.delete(old); }
      bind(doc, key);
      thumbLater(key);
    }, (e) => console.warn('recent files:', e));
  };
  const changes = new WeakMap();
  /** how many changes doc has had: a save leaves the document unchanged only if none came while it was written */
  VO.changeCount = (doc) => (doc && changes.get(doc)) || 0;
  /** doc has new unsaved changes: its unsaved version is written a little later */
  VO.changed = (doc) => {
    if (doc) changes.set(doc, (changes.get(doc) || 0) + 1);
    if (!current || !doc || !current.hooks.snapshot || lite) return;   // the lite edition keeps no unsaved versions
    const s = draftOf(doc), now = Date.now();
    if (s.since === null) s.since = now;
    clearTimeout(s.timer);
    const at = Math.min(Math.max(now + DRAFT_IDLE, s.last + s.gap), s.since + s.gap);
    s.timer = setTimeout(() => saveDraft(doc), Math.max(0, at - now));
  };
  async function saveDraft(doc) {
    const H = current.hooks, s = draftOf(doc);
    s.since = null;   // changes from here on wait for the next one
    if ((H.isDirty && !H.isDirty(doc)) || (H.autosave && H.autosave() === false)) return;
    s.last = Date.now();
    const t0 = performance.now();
    let blob = null;
    try { blob = await H.snapshot(doc); } catch (e) { console.warn('autosave:', e); return; }
    if (blob) s.gap = Math.max(DRAFT_GAP, DRAFT_PER_MB * blob.size / 1048576, (performance.now() - t0) * 30);
    if (!blob || (H.isDirty && !H.isDirty(doc))) return;
    const info = H.docInfo(doc);
    if (blob.size > MAX_BYTES) {
      if (!s.big && window.L && L.ui) L.ui.toast(VO.t('“{0}” is over 50 MB, so its changes aren\'t kept for recovery after a crash. Save it to keep them.', info.name), 10000);
      s.big = true;
      return;
    }
    if (native) {
      try { await setDraft(info, blob); } catch (e) { console.warn('autosave:', e); return; }
      thumbLater();
      return;
    }
    let key = docKeys.get(doc);
    if (!key) { key = current.app + ':~' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); bind(doc, key); }
    try { await VO.recent.setDraft(key, current.app, info.name, info.draftName, blob, info.lossState); } catch (e) { console.warn('autosave:', e); return; }
    if (!H.current || H.current() === doc) thumbLater(key);
  }
  /** the desktop app: this window's unsaved version, into its Recent Files entry; the body is what it says about
   *  itself (JSON, after its length in 4 bytes) and then its bytes (desktop/src/doc.rs, host_set_draft) */
  async function setDraft(info, blob) {
    const meta = new TextEncoder().encode(JSON.stringify({ name: info.name, draftName: info.draftName || undefined, lossState: info.lossState }));
    const body = new Uint8Array(4 + meta.length + blob.size);
    new DataView(body.buffer).setUint32(0, meta.length, true);
    body.set(meta, 4);
    body.set(await raw(blob), 4 + meta.length);
    return call('host_set_draft', body);
  }
  /** the user chose not to keep doc's changes (No in "Do you want to save the changes…?") */
  VO.discard = (doc) => {
    if (!doc) return;
    clearTimeout(draftOf(doc).timer);
    drafts.delete(doc);
    if (native) { if (!lite) call('host_clear_draft').catch(() => {}); return; }
    const key = docKeys.get(doc);
    if (!key) return;
    docKeys.delete(doc);
    openKeys.delete(key);
    VO.recent.discardDraft(key).catch(() => {});
  };

  /* ------------------------------------------------------------ the list inside the apps */
  const when = (t) => {
    const d = new Date(t), today = new Date().toDateString() === d.toDateString();
    const loc = VO.i18n ? VO.i18n.dates : [];
    return today ? d.toLocaleTimeString(loc, { hour: '2-digit', minute: '2-digit' }) : d.toLocaleDateString(loc, { month: 'short', day: 'numeric' });
  };
  /* its text in the UI's language (common/i18n.js) */
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = VO.t ? VO.t(text) : text; return e; };
  const icon = (name) => (window.L && L.icons ? L.icons.get(name) : '');
  /* the desktop app: Recent Files are the Start Center's alone; an app's lists are a way there */
  const toStart = (text) => { const b = el('button', 'tp-link', null); b.type = 'button'; const ic = el('span'); ic.innerHTML = icon('open'); b.append(ic, el('span', null, text)); b.addEventListener('click', () => VO.host.showRecent().catch(() => {})); return b; };
  /** File menu: the app's four most recent entries ({ label, run } items for L.ui menus); the desktop app's Recent
   *  Files... */
  VO.recentMenuItems = (max) => (lite ? [] : native ? [{ label: VO.t('Recent &Files...'), raw: true, run: () => VO.host.showRecent().catch(() => {}) }] : VO.recent.forApp().slice(0, max || 4).map((e, i) => ({
    label: '&' + (i + 1) + ' ' + (e.draft ? VO.t('{0} (unsaved changes)', e.name) : e.name).replace(/&/g, '&&'), raw: true,
    run: () => VO.openRecent(e.key, { draft: !!e.draft }),
  })));
  /** task pane: the app's recent entries as links (Getting Started ▸ Open); keeps itself up to date */
  VO.recentLinks = (max) => {
    const box = el('div', 'vo-recent');
    if (lite) return box;
    if (native) { box.append(toStart('Recent Files...')); return box; }
    const draw = () => {
      if (!box.isConnected && box.dataset.drawn) { stop(); return; }
      box.dataset.drawn = '1';
      box.replaceChildren(...VO.recent.forApp().slice(0, max || 4).map((e) => {
        const b = el('button', 'tp-link');
        b.type = 'button';
        b.title = e.draft ? VO.t('{0} — unsaved changes from {1}', e.name, new Date(e.draft.time).toLocaleString(VO.i18n ? VO.i18n.dates : [])) : e.name;
        b.translate = false;
        const ic = el('span'); ic.innerHTML = icon(e.draft ? 'save' : 'open');
        const name = el('span');
        name.textContent = e.draft ? VO.t('{0} (unsaved)', e.name) : e.name;
        b.append(ic, name);
        b.addEventListener('click', () => VO.openRecent(e.key, { draft: !!e.draft }));
        return b;
      }));
    };
    draw.cacheOnly = true;
    const stop = VO.onRecentChange(draw);
    draw();
    return box;
  };
  /** the Document Recovery task pane: this app's unsaved versions, each to open or to discard */
  VO.recoveryPane = (b) => {
    if (lite) { b.append(el('div', 'tp-note', 'VibeOffice Lite doesn\'t keep unsaved changes: save your document to keep them.')); return; }
    if (native) { b.append(el('div', 'tp-note', 'Unsaved changes are kept for each document in Recent Files, in the Start Center: open them from there, or discard them.'), toStart('Recent Files...')); return; }
    const list = el('div', 'vo-recovery');
    const draw = () => {
      if (!list.isConnected && list.dataset.drawn) { stop(); return; }
      list.dataset.drawn = '1';
      const drafts = VO.recent.forApp().filter((e) => e.draft);
      list.replaceChildren(...(drafts.length ? drafts.map((e) => {
        const item = el('div', 'tp-sec');
        const head = el('div', 'tp-h', e.name);
        const from = e.saved ? VO.t('Unsaved changes from {0}', when(e.draft.time)) : VO.t('Unsaved changes from {0} (never saved)', when(e.draft.time));
        const here = native ? e.here : openKeys.has(e.key);
        const note = el('div', 'tp-note', here ? VO.t('{0} — open in this window', from) : isOpen(e) ? VO.t('{0} — open in another window', from) : from);
        const open = el('button', 'tp-link'); open.type = 'button';
        const oi = el('span'); oi.innerHTML = icon('open'); open.append(oi, el('span', null, 'Open the unsaved version'));
        open.addEventListener('click', () => VO.openRecent(e.key, { draft: true }));
        const drop = el('button', 'tp-link'); drop.type = 'button';
        const di = el('span'); di.innerHTML = icon('delete') || icon('cut'); drop.append(di, el('span', null, 'Discard the changes'));
        drop.addEventListener('click', () => VO.recent.discardDraft(e.key));
        item.append(head, note, open, drop);
        return item;
      }) : [el('div', 'tp-note', 'No unsaved versions are waiting.')]));
    };
    draw.cacheOnly = true;
    const stop = VO.onRecentChange(draw);
    b.append(el('div', 'tp-note', VO.t('{0} kept the unsaved changes of these documents. Open the ones you want to keep and save them; discard the rest.', current ? VO.APPS[current.app].name : 'VibeOffice')), list);
    VO.recent.list().then((l) => { VO.recent.cache = l; draw(); }, () => draw());
  };
})();
