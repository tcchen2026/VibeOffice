// The desktop host (desktop/src/host.rs, doc.rs) as seen by the pages, for Chromium: run before each page's
// scripts with  node tools/shot.mjs '' steps.json --pre tests/desktop/host-stub.js
// This tab plays every new window the native side would make (it goes to that window's address). The disk holds
// FILES; the Open dialog "chooses" ?pick=<names> (comma-separated; none: Cancel). Recent Files start empty, or as
// ?recent=<names> (each a file's entry; a name ending in * has unsaved changes); ?doc pages hold ?file=<name> (or
// what the last window made was given). The pages have no browser storage, as in the app (host.rs, NO_STORAGE): the
// tab's sessionStorage is window.stubSession, for the stub and the tests' steps. Records there: hostCalls (each
// command, with the page), opened (host_open_url), uiScale, newWindow, closed (host_close_window), focused
// (host_focus), reopened (host_reopen's address), showRecent (host_show_recent), dialogs (the Open dialog's labels), chose (the Save dialog's arguments), saves (each write), touched
// (each use of a storage API the app removes); stubSettings, the settings (settings.rs, config.json).
// The clock says 2026-10-16 (?today=yyyy-mm-dd another day), so the versions here don't grow old (VO.isOutdated).
// ?web: no host, only the version a published copy of the site is stamped with.
// ?lite (on the first page; the tab keeps it): the lite edition (desktop/src/lite.rs): one window, whose pages all
// have its grants; its Documents hold FILES; records liteOpened, liteNew, liteHome, liteQuit.
(() => {
  const sessionStorage = window.stubSession = window.sessionStorage;
  const RealDate = Date;
  const shift = RealDate.parse(((/today=(\d{4}-\d{2}-\d{2})/.exec(location.search) || [])[1] || '2026-10-16') + 'T12:00') - RealDate.now();
  window.Date = class extends RealDate {
    constructor(...a) { if (a.length) super(...a); else super(RealDate.now() + shift); }
    static now() { return RealDate.now() + shift; }
  };
  if (/[?&]web\b/.test(location.search)) {   // links open in a new tab: recorded instead
    window.__VO_HOST__ = { appVersion: (/version=(v[0-9.]+-[0-9a-f]{6})/.exec(location.search) || [])[1] || 'v2026.10.11-abc123' };
    window.open = (url) => { sessionStorage.opened = url; return null; };
    return;
  }
  const log = (k) => { sessionStorage[k] = (sessionStorage[k] ? sessionStorage[k] + ' ' : '') + location.pathname; };
  // ?update=button / ?update=dialog: updater/current.json asks this version to update; ?update=none asks nothing of
  // any version since January; ?update=offline: it cannot be read; otherwise the site's own file
  const ask = (/update=(button|dialog|none|offline)/.exec(location.search) || [])[1];
  const newer = { web: 'v2026.10.12-000000', macos: 'v2026.10.12-000000', windows: 'v2026.10.12-000000', linux: 'v2026.10.12-000000' };
  const same = { web: 'v2026.10.11-000000', macos: 'v2026.10.11-000000', windows: 'v2026.10.11-000000', linux: 'v2026.10.11-000000' };
  const news = { 'v2026.10.10-000000': 'Older.', 'v2026.10.12-000000': 'Twelfth.\nSecond line.', 'v2026.10.11-def456': 'Same day.', 'v2026.10.14-000000': 'Fourteenth.' };
  const json = (o) => 'data:application/json,' + encodeURIComponent(JSON.stringify(o));
  const january = { web: 'v2026.01.01-000000', macos: 'v2026.01.01-000000', windows: 'v2026.01.01-000000', linux: 'v2026.01.01-000000' };
  const policy = { button: { 'show-update-button': newer, 'show-update-dialog': same }, dialog: { 'show-update-button': newer, 'show-update-dialog': newer },
    none: { 'show-update-button': january, 'show-update-dialog': january } }[ask];
  const updates = ask === 'offline' ? 'data:,not JSON' : policy && json(policy);
  // ?version=v…: the build's version (default v2026.10.11-abc123)
  const appVersion = (/version=(v[0-9.]+-[0-9a-f]{6})/.exec(location.search) || [])[1] || 'v2026.10.11-abc123';
  // ?appstore: the Mac App Store edition, which the store updates (host.rs, init_script)
  const appStore = /appstore/.test(location.search);
  if (/[?&]lite\b/.test(location.search)) sessionStorage.stubLite = '1';
  const lite = !!sessionStorage.stubLite;
  // ?sandboxed: the Mac App Sandbox, where a refused save may ask for the file's folder (host_allow_access)
  window.__VO_HOST__ = Object.freeze({ platform: lite ? 'linux' : appStore ? 'macos' : 'windows', version: 1, appVersion, edition: lite ? 'lite' : appStore ? 'appstore' : 'direct',
    sandboxed: /[?&]sandboxed\b/.test(location.search),
    updates: updates || '/updater/current.json', whatsnew: ask ? json(news) : '/updater/whatsnew.json' });
  // the commands each kind of window may call (desktop/permissions/host.toml, desktop/capabilities/); the licences
  // page none
  const kind = lite ? 'lite' : /^\/(index\.html)?$/.test(location.pathname) ? 'start' : /^\/(quire|ledger|lectern)\//.test(location.pathname) ? 'app' : 'none';
  const app = /^\/(quire|ledger|lectern)\//.test(location.pathname) ? location.pathname.split('/')[1] : null;
  const grants = {
    start: ['host_open', 'host_open_recent', 'host_new_document', 'host_recent', 'host_recent_preview', 'host_recent_remove',
      'host_recent_clear', 'host_discard_draft', 'host_new_window', 'host_open_url', 'host_set_setting'],
    app: ['host_document', 'host_document_bytes', 'host_open', 'host_reopen', 'host_show_recent', 'host_new_document',
      'host_set_preview', 'host_set_draft', 'host_clear_draft', 'host_windows',
      'host_focus', 'host_closing', 'host_save', 'host_allow_access', 'host_choose', 'host_write_chosen', 'host_cancel_choice', 'host_versions',
      'host_open_version', 'host_new_window', 'host_close_window', 'host_set_setting'],
    lite: ['host_document', 'host_document_bytes', 'host_open', 'host_reopen', 'host_show_recent', 'host_new_document',
      'host_set_preview', 'host_set_draft', 'host_clear_draft', 'host_windows',
      'host_focus', 'host_closing', 'host_save', 'host_allow_access', 'host_choose', 'host_write_chosen', 'host_cancel_choice', 'host_versions',
      'host_open_version', 'host_new_window', 'host_close_window', 'host_set_setting',
      'host_lite_list', 'host_lite_open', 'host_lite_new', 'host_lite_home', 'host_lite_licenses', 'host_lite_quit',
      'host_lite_show'],
    none: [],
  }[kind];
  /* the disk, the apps' types (desktop/src/types.rs) and Recent Files (doc.rs, store.rs) */
  const FILES = { 'Budget.csv': 'Item,Cost\nRent,900\nFood,250\n', 'Other.csv': 'A,B\n1,2\n', 'Notes.md': '# Notes\n\nHello', 'Report.txt': 'Hello there' };
  const appFor = (name) => ({ csv: 'ledger', md: 'quire', txt: 'quire', docx: 'quire', pptx: 'lectern' })[name.split('.').pop()];
  const q = new URLSearchParams(location.search);
  const recent = () => JSON.parse(sessionStorage.stubRecent || 'null') || (q.get('recent') || '').split(',').filter(Boolean).map((n, i) => {
    const name = n.replace(/\*$/, '');
    return { key: 'k-' + name, app: appFor(name), name, folder: 'C:\\Users\\me\\Documents', time: 1000 - i, size: 28, saved: true, preview: true,
      draft: n.endsWith('*') ? { name, time: 1000 - i } : null, open: false, here: false };
  });
  const keep = (list) => { sessionStorage.stubRecent = JSON.stringify(list); window.dispatchEvent(new Event('vo-recent')); };
  const bytes = (text) => new TextEncoder().encode(text).buffer;
  /* a new window for a document: this tab goes there, holding it */
  const windowFor = (doc) => {
    sessionStorage.stubDoc = JSON.stringify(doc);
    sessionStorage.newWindow = `${location.origin}/${appFor(doc.info.fileName) || app}/?doc`;
    setTimeout(() => { location.href = sessionStorage.newWindow; });
  };
  const fileDoc = (name, draft) => ({ key: 'k-' + name, info: { fileName: name, name, saved: true, draft: !!draft, lossState: null, bound: true, readOnly: q.has('readonly') },
    text: draft ? FILES[name].replace(/\n$/, '') + '\nDraft,1\n' : FILES[name] });
  if (q.has('doc') && !sessionStorage.stubDoc && q.get('file')) sessionStorage.stubDoc = JSON.stringify(fileDoc(q.get('file')));
  /* opens a file or entry as doc.rs does: always in a new window, never into the page asking (Docs::adopt) */
  const open = (doc) => {
    const list = recent().filter((e) => e.key !== doc.key);
    list.unshift({ key: doc.key, app: appFor(doc.info.name), name: doc.info.name, folder: 'C:\\Users\\me\\Documents', time: Date.now(), size: 28, saved: true, preview: false, draft: null, open: true, here: false });
    keep(list);
    windowFor(doc);
  };
  const text = (body) => new TextDecoder().decode(body);
  window.__TAURI_INTERNALS__ = {
    async invoke(cmd, args, options) {
      log('hostCalls');
      sessionStorage.hostCalls += ':' + cmd;
      if (!grants.includes(cmd)) throw `Command ${cmd} not allowed by ACL`;
      if (lite) {   // lite.rs: the one window goes to the page; the documents are FILES, newest first
        const go = (url) => { sessionStorage.liteAt = url; setTimeout(() => { location.href = url; }); return null; };
        if (cmd === 'host_lite_list') {
          return { documents: Object.keys(FILES).map((name, i) => ({ name, app: appFor(name), size: FILES[name].length, modified: Date.now() - i * 86400000 })),
            folder: '~/VibeOffice/Documents', sftp: ['pi@192.168.1.23'] };
        }
        if (cmd === 'host_lite_open') {
          if (!FILES[args.name]) throw `“${args.name}” is no longer in Documents`;
          const token = Math.random().toString(36).slice(2);
          sessionStorage.stubDoc = JSON.stringify({ ...fileDoc(args.name), token });
          sessionStorage.liteOpened = args.name;
          return go(`/${appFor(args.name)}/?doc&t=${token}`);
        }
        if (cmd === 'host_lite_new') {
          if (!['quire', 'ledger', 'lectern'].includes(args.page)) throw `not one of the apps: ${args.page}`;
          sessionStorage.liteNew = args.page;
          delete sessionStorage.stubKey;
          return go(`/${args.page}/?` + (args.start && args.start.template ? 'template=' + args.start.template : 'new') + '&n=1');
        }
        if (cmd === 'host_lite_home') { log('liteHome'); delete sessionStorage.stubKey; return go('/lite.html'); }
        if (cmd === 'host_lite_licenses') return go('/oss-licenses.html');
        if (cmd === 'host_lite_quit') { sessionStorage.liteQuit = '1'; return null; }
        if (cmd === 'host_lite_show') { log(args.on ? 'liteShow' : 'liteShown'); return false; }   // no HDMI screen here
        if (cmd === 'host_new_window') throw 'the lite edition has one window';
        if (cmd === 'host_close_window') return false;
        if (cmd === 'host_choose') {   // no dialog: the page's name, the type's extension, in Documents (doc.rs)
          sessionStorage.chose = JSON.stringify(args);
          const ext = { html: 'htm', mdzip: 'zip', xmlss: 'xml' }[args.type] || args.type, name = args.name + '.' + ext;
          if (FILES[name] && !args.replace) throw { code: 'exists', message: `“${name}” already exists.`, file: name };
          sessionStorage.choice = JSON.stringify({ name, type: args.type, saveAs: args.purpose === 'saveAs' });
          return { name, type: args.type };
        }
      }
      if (cmd === 'host_open') {   // an app window: its app's first file is held for host_reopen (next)
        sessionStorage.dialogs = JSON.stringify(args.labels);
        let opened = 0, here = null;
        for (const name of (q.get('pick') || '').split(',').filter(Boolean)) {
          if (!here && kind === 'app' && appFor(name) === app) { here = name; sessionStorage.next = name; } else { open(fileDoc(name)); opened++; }
        }
        return { opened, here };
      }
      if (cmd === 'host_reopen') {   // this tab reloads to the app's page with a one-time word (token)
        const name = sessionStorage.next;
        delete sessionStorage.next;
        if (!args.go || !name) return null;
        const token = Math.random().toString(36).slice(2);
        sessionStorage.stubDoc = JSON.stringify({ ...fileDoc(name), token });
        delete sessionStorage.stubKey;
        sessionStorage.reopened = `/${app}/?doc&t=${token}`;
        setTimeout(() => { location.href = sessionStorage.reopened; });
        return null;
      }
      if (cmd === 'host_show_recent') { log('showRecent'); return null; }
      if (cmd === 'host_open_recent') {
        const e = recent().find((x) => x.key === args.key);
        if (!e) throw 'That file is no longer in Recent Files.';
        open(fileDoc(e.name, args.draft));
        return null;
      }
      if (cmd === 'host_new_document') {   // raw bytes, the name in X-Name
        const name = decodeURIComponent(options.headers['X-Name']);
        if (!(args instanceof Uint8Array)) throw 'no bytes';
        windowFor({ key: '', info: { fileName: name, name, saved: !options.headers['X-Made'], draft: false, lossState: null, bound: false, readOnly: false }, text: text(args) });
        return null;
      }
      // a window reloaded for a file gives it only to the page with its word (doc.rs, takes)
      const mine = (d) => d && (!d.token || (args && args.t) === d.token) ? d : null;
      if (cmd === 'host_document') {
        const d = mine(JSON.parse(sessionStorage.stubDoc || 'null'));
        if (d && d.key) sessionStorage.stubKey = d.key;
        return d && d.info;
      }
      if (cmd === 'host_document_bytes') {
        const d = mine(JSON.parse(sessionStorage.stubDoc || 'null'));
        if (!d) throw 'no document';
        delete sessionStorage.stubDoc;
        // b64: bytes that aren't text (a workbook's draft, as Ledger's snapshot makes it)
        return d.b64 ? Uint8Array.from(atob(d.b64), (c) => c.charCodeAt(0)).buffer : bytes(d.text);
      }
      if (cmd === 'host_recent') return recent().map((e) => ({ ...e, open: e.key === sessionStorage.stubKey, here: e.key === sessionStorage.stubKey }));
      if (cmd === 'host_recent_preview') return bytes('\xff\xd8\xff');
      if (cmd === 'host_recent_remove' || cmd === 'host_discard_draft') {
        if (args.key === sessionStorage.stubKey) throw 'That document is open in a window.';
        keep(cmd === 'host_recent_remove' ? recent().filter((e) => e.key !== args.key) : recent().map((e) => (e.key === args.key ? { ...e, draft: null } : e)));
        return null;
      }
      if (cmd === 'host_recent_clear') { keep(recent().filter((e) => e.draft)); return null; }
      if (cmd === 'host_set_preview' || cmd === 'host_clear_draft') return null;
      if (cmd === 'host_set_draft') {   // its length (4 bytes), what it says about itself (JSON), its bytes
        const n = new DataView(args.buffer).getUint32(0, true);
        sessionStorage.draft = JSON.stringify({ ...JSON.parse(text(args.subarray(4, 4 + n))), size: args.length - 4 - n });
        return null;
      }
      if (cmd === 'host_windows') return [{ id: app + '-1', title: document.title, current: true }, { id: app + '-2', title: 'Other - ' + app, current: false }];
      if (cmd === 'host_focus') { sessionStorage.focused = args.id; return null; }
      if (cmd === 'host_closing') return null;
      /* saving (doc.rs, files.rs): ?fail=<code> fails the next write once; the Save dialog "picks" sessionStorage
         stubSave (a name, or none: Cancel); writes are recorded in sessionStorage.saves */
      const wrote = (what) => { const all = JSON.parse(sessionStorage.saves || '[]'); all.push(what); sessionStorage.saves = JSON.stringify(all); };
      const failing = () => { const f = q.get('fail'); if (f && !sessionStorage.failed) { sessionStorage.failed = f; throw { code: f, message: 'stub ' + f }; } };
      if (cmd === 'host_save') {
        if (!sessionStorage.stubKey) throw { code: 'missing', message: 'This document has no file of its own yet.' };
        failing();
        wrote({ cmd, size: args.length, replace: !!(options && options.headers && options.headers['X-Replace']), head: text(args.subarray(0, 2)) });
        return null;
      }
      if (cmd === 'host_choose') {
        sessionStorage.chose = JSON.stringify(args);
        const name = sessionStorage.stubSave;
        if (!name) return null;
        const type = args.types.find((t) => ({ htm: 'html', zip: 'mdzip', xml: 'xmlss', markdown: 'md' }[name.split('.').pop()] || name.split('.').pop()) === t.id);
        sessionStorage.choice = JSON.stringify({ name, type: (type || args.types[0]).id, saveAs: args.purpose === 'saveAs' });
        return { name, type: (type || args.types[0]).id };
      }
      if (cmd === 'host_write_chosen') {
        const c = JSON.parse(sessionStorage.choice || 'null');
        if (!c) throw { code: 'io', message: 'nothing was chosen' };
        failing();
        delete sessionStorage.choice;
        const bound = c.saveAs && ['docx', 'docm', 'dotx', 'dotm', 'md', 'mdzip', 'xlsx', 'xlsm', 'xltx', 'xltm', 'xmlss', 'pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm'].includes(c.type);
        if (bound) sessionStorage.stubKey = 'k-' + c.name;
        wrote({ cmd, name: c.name, type: c.type, size: args.length, bound });
        return { name: c.name, type: c.type, bound };
      }
      if (cmd === 'host_cancel_choice') { delete sessionStorage.choice; return null; }
      if (cmd === 'host_allow_access') {   // the folder chooser "picks" the file's folder (recorded: allowed)
        if (!__VO_HOST__.sandboxed) throw { code: 'io', message: 'Only the Mac\'s sandbox asks for folders.' };
        sessionStorage.allowed = JSON.stringify(args);
        return true;
      }
      if (cmd === 'host_versions') return sessionStorage.stubKey ? [{ id: sessionStorage.stubKey + '.1', time: 1792230000000, size: 2048 }] : [];
      if (cmd === 'host_open_version') { sessionStorage.openedVersion = args.id; return null; }
      if (cmd === 'host_open_url') {
        // host::is_ours: the update page and our GitHub
        if (!/^https:\/\/(vibeoffice\.work\/updater\/|github\.com\/tcchen2026\/VibeOffice([/?#]|$))/.test(args.url)) throw `not opened: ${args.url}`;
        sessionStorage.opened = args.url;
        return null;
      }
      if (cmd === 'host_set_setting') {   // settings.rs: the key's form, the UI size's range
        const { key, value } = args;
        if (!/^[a-z][a-z0-9]*\.[A-Za-z0-9_.-]{1,64}$/.test(key)) throw `not a setting: ${key}`;
        if (key === 'vibeoffice.uiScale' && value != null && !(+value >= 0.5 && +value <= 4)) throw 'the UI size is from 0.5 to 4';
        const all = JSON.parse(sessionStorage.stubSettings || '{}');
        if (value == null) delete all[key]; else all[key] = value;
        sessionStorage.stubSettings = JSON.stringify(all);
        if (key === 'vibeoffice.uiScale') sessionStorage.uiScale = value;
        return null;
      }
      if (cmd === 'host_new_window') {   // this tab plays the new window, at the address host.rs makes (page_path)
        const { page, start = {} } = args;
        if (!['quire', 'ledger', 'lectern', 'licenses'].includes(page)) throw `not one of the suite's pages: ${page}`;
        if (kind !== 'start' && page !== 'licenses' && page !== app) throw `${app} doesn't open ${page}`;
        const url = new URL(page === 'licenses' ? '/oss-licenses.html' : `/${page}/`, location.origin);
        for (const [k, v] of Object.entries(start)) {
          if (!['new', 'template'].includes(k)) throw `unknown field \`${k}\``;
          url.searchParams.append(k, v === true ? '' : v);
        }
        sessionStorage.newWindow = url.href;
        setTimeout(() => { location.href = url.href; });
        return null;
      }
      if (cmd === 'host_close_window') {   // app pages are app windows; the Start Center is "main"
        if (location.pathname === '/') return false;
        sessionStorage.closed = location.pathname;
        return true;
      }
      throw `command ${cmd} not found`;
    },
  };
  // the settings as they are now, at the top of common/look.js (settings.rs, Served)
  // (the lite edition's defaults: settings.rs, LITE)
  const LITE = { 'vibeoffice.look': 'classic', 'vibeoffice.uiScale': '0.75',
    'quire.opts': '{"taskOpen":false,"startupPane":false,"ruler":false,"vruler":false,"view":"web","spell":false,"toolbars":{"standard":false,"formatting":false,"drawing":false,"tables":false,"reviewing":false,"picture":false,"outlining":false,"headerFooter":false,"mailMerge":false,"wordart":false}}',
    'ledger.opts': '{"taskOpen":false,"toolbars":{"standard":false,"formatting":false}}',
    'lectern.opts': '{"taskOpen":false,"startupPane":false,"layoutPaneOnNew":false,"leftOpen":false,"notesH":24,"spell":false,"toolbars":{"standard":false,"formatting":false,"drawing":false}}' };
  window.__VO_SETTINGS__ = Object.assign(lite ? LITE : { 'vibeoffice.look': 'paper' }, JSON.parse(sessionStorage.stubSettings || '{}'));
  // no browser storage (host.rs, NO_STORAGE): a use is recorded, and throws
  for (const [o, k] of [[window, 'localStorage'], [window, 'sessionStorage'], [window, 'indexedDB'], [window, 'caches'], [window, 'BroadcastChannel'], [navigator, 'storage'], [navigator, 'serviceWorker']]) {
    Object.defineProperty(o, k, { get() { log('touched'); sessionStorage.touched += ':' + k; throw new Error(k + ' is not used in the desktop app'); }, configurable: false });
  }
})();
