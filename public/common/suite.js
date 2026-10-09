/* VibeOffice suite services, shared by the start page and the apps (classic script, window.VO).
 *
 * - Which app opens which file (VO.APPS, VO.appFor).
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
  const VO = (window.VO = {});
  VO.VERSION = 'dev';   // the published copy says YYYY.MM.<commit> (stamped by the sync script)
  const BASE = new URL('..', document.currentScript.src).href;   // public/
  VO.base = BASE;

  /* ------------------------------------------------------------ apps */
  VO.APPS = {
    quire: { name: 'Quire', kind: 'Document', ext: ['docx', 'docm', 'dotx', 'dotm', 'doc', 'rtf', 'htm', 'html', 'txt', 'md', 'markdown', 'zip'] },
    ledger: { name: 'Ledger', kind: 'Spreadsheet', ext: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xls', 'csv', 'tsv', 'prn', 'xml'] },
    lectern: { name: 'Lectern', kind: 'Presentation', ext: ['pptx', 'pptm', 'ppsx', 'ppsm', 'potx', 'potm', 'ppt'] },
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

  /* ------------------------------------------------------------ service worker */
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
    navigator.serviceWorker.register(BASE + 'sw.js').catch((e) => console.warn('service worker:', e.message));
  }

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
  const chan = 'BroadcastChannel' in window ? new BroadcastChannel('vibeoffice') : null;
  const listeners = new Set();
  const notify = () => { VO.recent.refresh(); for (const fn of listeners) if (!fn.cacheOnly) fn(); };   // cache users redraw after the refresh
  const changed = () => { if (chan) chan.postMessage('recent'); notify(); };
  if (chan) chan.addEventListener('message', (e) => { if (e.data === 'recent') notify(); });
  /** call fn when the Recent Files list changes (here or in another page); returns the function that stops it */
  VO.onRecentChange = (fn) => { listeners.add(fn); return () => listeners.delete(fn); };

  /* ------------------------------------------------------------ Recent Files */
  const MAX_RECENT = 30;
  const MAX_BYTES = 50 * 1024 * 1024;   // bigger files are not kept (open them from disk again)
  const prune = async () => {
    /* entries with unsaved changes are never dropped */
    const all = (await VO.recent.list()).filter((e) => !e.draft);
    if (all.length > MAX_RECENT) await tx('recent', 'readwrite', (s) => { for (const e of all.slice(MAX_RECENT)) s.delete(e.key); });
  };
  VO.recent = {
    /** newest first: { key, app, name, size, time, file: Blob|null (opened or saved), thumb: Blob|null,
     *  draft: { file: Blob, name, time } | null (unsaved changes) }; a never-saved document has only a draft */
    async list() {
      const all = (await tx('recent', 'readonly', (s) => s.getAll())) || [];
      return all.sort((a, b) => b.time - a.time);
    },
    get: (key) => tx('recent', 'readonly', (s) => s.get(key)),
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
    },
    /** throw the unsaved version away; a document that was never saved leaves the list */
    async discardDraft(key) {
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
    async remove(key) { await tx('recent', 'readwrite', (s) => s.delete(key)); changed(); },
    async clear() { await tx('recent', 'readwrite', (s) => s.clear()); changed(); },
  };

  /* ------------------------------------------------------------ handing files to apps */
  /** where an app page opens: a new window when installed (the start page stays, as in LibreOffice), else this tab */
  VO.go = (url) => {
    const standalone = matchMedia('(display-mode: standalone)').matches || matchMedia('(display-mode: window-controls-overlay)').matches;
    if (standalone && window.open(url, '_blank')) return;
    location.href = url;
  };
  /** open a file in the app made for it; false if no app opens this kind of file */
  VO.handoff = async (file) => {
    const app = VO.appFor(file.name);
    if (!app) return false;
    const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    const day = Date.now() - 86400000;
    await tx('handoff', 'readwrite', (s) => {
      s.put({ id, name: file.name, file, time: Date.now() });
      s.openCursor().onsuccess = (e) => { const c = e.target.result; if (!c) return; if (c.value.time < day) c.delete(); c.continue(); };
    });
    VO.go(VO.url(app, { open: id }));
    return true;
  };

  /* ------------------------------------------------------------ the app side */
  const q = new URLSearchParams(location.search);
  /** how this app page was opened: null (plain: show the sample), { new }, { template: id }, { open: id } or { recent: key } */
  VO.launch = q.has('open') ? { open: q.get('open') } : q.has('recent') ? { recent: q.get('recent'), draft: q.has('draft') }
    : q.has('template') ? { template: q.get('template') } : q.has('new') ? { new: true } : null;

  let current = null;   // { app, hooks }
  /** an app is ready: open the file it was launched with, if any, then show the Document Recovery pane when
   *  other documents of this app have unsaved changes.
   *  hooks: open(file, { draft, name, saved }), thumb() -> a canvas (or image) of the first page, message(text),
   *  templates() -> [{ id, name }] (read by tools/templates.mjs); for autosave: snapshot(doc) -> Blob in the
   *  app's format, docInfo(doc) -> { name, draftName }, isDirty(doc), current() -> doc, autosave() -> false
   *  when turned off; showRecovery() opens the Document Recovery pane */
  VO.attach = async (app, hooks) => {
    current = { app, hooks: hooks || {} };
    VO.recent.refresh();
    const l = VO.launch;
    if (l && !l.new) {
      if (history.replaceState) history.replaceState(null, '', location.pathname);   // a reload starts clean
      if (l.open) {
        let rec = null;
        try {
          rec = await tx('handoff', 'readonly', (s) => s.get(l.open));
          if (rec) await tx('handoff', 'readwrite', (s) => s.delete(l.open));
        } catch (e) { console.error(e); }
        if (rec && rec.file) await current.hooks.open(asFile(rec.file, rec.name), {});
        else (current.hooks.message || alert)('That file is no longer available here. Open it again from your computer.');
      } else if (l.recent) await VO.openRecent(l.recent, { draft: l.draft });
    }
    setTimeout(async () => {
      const others = (await VO.recent.list()).filter((e) => e.app === app && e.draft && !openKeys.has(e.key));
      if (others.length && current.hooks.showRecovery) current.hooks.showRecovery();
    }, 700);
  };
  const asFile = (blob, name) => (blob instanceof File && blob.name === name ? blob : new File([blob], name, { type: blob.type }));
  /** open a Recent Files entry in this app: its unsaved version when draft is true, else the saved file */
  VO.openRecent = async (key, o) => {
    let e = null;
    try { e = await VO.recent.get(key); } catch (err) { console.error(err); }
    const draft = !!(o && o.draft && e && e.draft);
    const blob = e && (draft ? e.draft.file : e.file);
    if (!blob) { (current.hooks.message || alert)('That file is no longer available here. Open it again from your computer.'); return; }
    const f = asFile(blob, draft ? e.draft.name : e.name);
    if (draft) f.voDraft = key;
    await current.hooks.open(f, { draft, name: e.name, saved: !!e.file, lossState: draft ? e.draft.lossState : undefined });
  };
  /** File ▸ New… inside an app: Blank and the app's templates with their previews (the Start Center's
   *  Create File gallery, common/templates.json), in the app's own dialog. create(id): id '' is Blank. */
  VO.newDialog = async (app, create) => {
    const ui = window.L && window.L.ui;
    if (!ui || !ui.dialog) return create('');
    let list = null;
    try { list = (await (await fetch(BASE + 'common/templates.json')).json())[app]; } catch (e) { console.warn('templates:', e); }
    if (!list || !list.length) list = [{ id: '', name: 'Blank', img: null }];
    if (!document.getElementById('vo-gallery-css')) {
      const st = document.createElement('style');
      st.id = 'vo-gallery-css';
      st.textContent = '.vo-gal{display:grid;grid-template-columns:repeat(auto-fill,minmax(132px,1fr));gap:8px;max-height:min(62vh,540px);overflow:auto;padding:6px;background:#fff;border:1px solid #7F9DB9}'
        + '.vo-tpl{display:flex;flex-direction:column;gap:4px;padding:5px;border:1px solid transparent;border-radius:2px;background:none;text-align:left;cursor:default;font:inherit;color:#000}'
        + '.vo-tpl:hover,.vo-tpl:focus-visible{background:#FFEEC2;border-color:#E3A23B;outline:none}'
        + '.vo-tpl span{display:flex;align-items:center;justify-content:center;height:112px;background:#f3f6fb;border:1px solid #b8c7df;overflow:hidden}'
        + '.vo-tpl img{max-width:100%;max-height:100%;box-shadow:0 1px 2px rgba(0,0,0,.25)}'
        + '.vo-tpl b{font-weight:normal;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}'
        + '.vo-gal-note{margin:0 0 6px}';
      document.head.appendChild(st);
    }
    const gal = document.createElement('div');
    gal.className = 'vo-gal';
    let d = null;
    for (const t of list) {
      const b = document.createElement('button');
      b.type = 'button'; b.className = 'vo-tpl'; b.title = t.name;
      const pic = document.createElement('span');
      if (t.img) { const img = document.createElement('img'); img.src = BASE + t.img; img.alt = ''; pic.append(img); }
      const name = document.createElement('b');
      name.textContent = t.name;
      b.append(pic, name);
      b.addEventListener('click', () => { if (d) d.close(-1); create(t.id); });
      gal.append(b);
    }
    const body = document.createElement('div');
    const note = document.createElement('p');
    note.className = 'vo-gal-note';
    note.textContent = 'Choose Blank or a template:';
    body.append(note, gal);
    d = ui.dialog({ title: 'New ' + VO.APPS[app].kind, width: 700, body, buttons: [{ label: 'Cancel', cancel: true }], focus: gal.firstChild });
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
  /** the thumbnail for a Recent Files entry, a little after the document has been drawn */
  function thumbLater(key) {
    if (!key) return;
    setTimeout(async () => {
      try { const blob = await VO.thumbnail(); if (blob) await VO.recent.setThumb(key, blob); }
      catch (e) { console.warn('thumbnail:', e); }
    }, 600);
  }
  /* ------------------------------------------------------------ documents, autosave */
  const docKeys = new WeakMap();   // the app's document object -> its Recent Files key
  const openKeys = new Set();      // keys of the documents open in this page (left out of Document Recovery)
  const timers = new WeakMap(), slow = new WeakMap();
  const DRAFT_DELAY = 15000;
  const bind = (doc, key) => { if (doc) docKeys.set(doc, key); openKeys.add(key); };
  /** the app opened a file (from any source) as doc */
  VO.opened = (file, doc) => {
    if (!current) return;
    if (file.voDraft) { bind(doc, file.voDraft); return; }   // an unsaved version: keep its entry as it is
    VO.recent.put(current.app, file.name, file, true).then((key) => { if (key) bind(doc, key); thumbLater(key); }, (e) => console.warn('recent files:', e));
  };
  /** the app saved doc in its own format as name: that is its entry now, without unsaved changes */
  VO.saved = (name, blob, doc) => {
    if (!current) return;
    const old = doc && docKeys.get(doc);
    if (doc) clearTimeout(timers.get(doc));
    VO.recent.put(current.app, name, blob, false).then(async (key) => {
      if (!key) return;
      if (old && old !== key) { await VO.recent.discardDraft(old); openKeys.delete(old); }
      bind(doc, key);
      thumbLater(key);
    }, (e) => console.warn('recent files:', e));
  };
  /** doc has new unsaved changes: its unsaved version is written a little later */
  VO.changed = (doc) => {
    if (!current || !doc || !current.hooks.snapshot) return;
    clearTimeout(timers.get(doc));
    timers.set(doc, setTimeout(() => saveDraft(doc), Math.max(DRAFT_DELAY, slow.get(doc) || 0)));
  };
  async function saveDraft(doc) {
    const H = current.hooks;
    if ((H.isDirty && !H.isDirty(doc)) || (H.autosave && H.autosave() === false)) return;
    const t0 = performance.now();
    let blob = null;
    try { blob = await H.snapshot(doc); } catch (e) { console.warn('autosave:', e); return; }
    slow.set(doc, (performance.now() - t0) * 30);   // a document that takes 1 s to write is saved every 30 s at most
    if (!blob || blob.size > MAX_BYTES || (H.isDirty && !H.isDirty(doc))) return;
    const info = H.docInfo(doc);
    let key = docKeys.get(doc);
    if (!key) { key = current.app + ':~' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); bind(doc, key); }
    try { await VO.recent.setDraft(key, current.app, info.name, info.draftName, blob, info.lossState); } catch (e) { console.warn('autosave:', e); return; }
    if (!H.current || H.current() === doc) thumbLater(key);
  }
  /** the user chose not to keep doc's changes (No in "Do you want to save the changes…?") */
  VO.discard = (doc) => {
    if (!doc) return;
    clearTimeout(timers.get(doc));
    const key = docKeys.get(doc);
    if (!key) return;
    docKeys.delete(doc);
    openKeys.delete(key);
    VO.recent.discardDraft(key).catch(() => {});
  };

  /* ------------------------------------------------------------ the list inside the apps */
  const when = (t) => {
    const d = new Date(t), today = new Date().toDateString() === d.toDateString();
    return today ? d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : d.toLocaleDateString([], { month: 'short', day: 'numeric' });
  };
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const icon = (name) => (window.L && L.icons ? L.icons.get(name) : '');
  /** File menu: the app's four most recent entries ({ label, run } items for L.ui menus) */
  VO.recentMenuItems = (max) => VO.recent.forApp().slice(0, max || 4).map((e, i) => ({
    label: `&${i + 1} ${e.name.replace(/&/g, '&&')}${e.draft ? ' (unsaved changes)' : ''}`,
    run: () => VO.openRecent(e.key, { draft: !!e.draft }),
  }));
  /** task pane: the app's recent entries as links (Getting Started ▸ Open); keeps itself up to date */
  VO.recentLinks = (max) => {
    const box = el('div', 'vo-recent');
    const draw = () => {
      if (!box.isConnected && box.dataset.drawn) { stop(); return; }
      box.dataset.drawn = '1';
      box.replaceChildren(...VO.recent.forApp().slice(0, max || 4).map((e) => {
        const b = el('button', 'tp-link');
        b.type = 'button';
        b.title = e.draft ? `${e.name} — unsaved changes from ${new Date(e.draft.time).toLocaleString()}` : e.name;
        const ic = el('span'); ic.innerHTML = icon(e.draft ? 'save' : 'open');
        b.append(ic, el('span', null, e.name + (e.draft ? ' (unsaved)' : '')));
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
    const list = el('div', 'vo-recovery');
    const draw = () => {
      if (!list.isConnected && list.dataset.drawn) { stop(); return; }
      list.dataset.drawn = '1';
      const drafts = VO.recent.forApp().filter((e) => e.draft);
      list.replaceChildren(...(drafts.length ? drafts.map((e) => {
        const item = el('div', 'tp-sec');
        const head = el('div', 'tp-h', e.name);
        const note = el('div', 'tp-note', `Unsaved changes from ${when(e.draft.time)}` + (e.file ? '' : ' (never saved)') + (openKeys.has(e.key) ? ' — open in this window' : ''));
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
    b.append(el('div', 'tp-note', `${current ? VO.APPS[current.app].name : 'VibeOffice'} kept the unsaved changes of these documents. Open the ones you want to keep and save them; discard the rest.`), list);
    VO.recent.list().then((l) => { VO.recent.cache = l; draw(); }, () => draw());
  };
})();
